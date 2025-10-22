import glob
import time
import datetime
import logging
import os
from pathlib import Path
from contextlib import contextmanager
from typing import List, Tuple

import loginInfo
from RecaptchaSolver import RecaptchaSolver

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, WebDriverException


class InvoiceDownloader:
    """Automated invoice downloader for Taiwan e-invoice system."""
    
    def __init__(self, webdriver_path: str = r"chromedriver-win64\chromedriver.exe", 
                 download_dir: str = rf"C:\Users\{loginInfo.User}\Downloads",
                 prefix: str = f"{loginInfo.ban}_IN_{datetime.datetime.today().strftime('%Y%m%d')}",
                 timeout: int = 30,
                 recaptcha_solver=None):
        self.webdriver_path = webdriver_path
        self.download_dir = Path(download_dir)
        self.timeout = timeout
        self.browser = None
        self.total_downloads = 0
        self.prefix = prefix
        self.pattern = str(self.download_dir / f"{self.prefix}*.xls")
        self.recaptcha_solver = recaptcha_solver
        
        # Setup logging
        logging.basicConfig(level=logging.INFO,
                            filename='einvoice.log',
                            format='%(asctime)s - %(levelname)s - %(message)s',
                            encoding='utf-8')
        self.logger = logging.getLogger(__name__)
    
    @contextmanager
    def get_browser(self):
        """Context manager for browser lifecycle management."""
        try:
            service = Service(self.webdriver_path)
            options = self._get_chrome_options()
            self.browser = webdriver.Chrome(service=service, options=options)
            self.browser.implicitly_wait(self.timeout)
            yield self.browser
        except WebDriverException as e:
            self.logger.error(f"Failed to initialize browser: {e}")
            raise
        finally:
            if self.browser:
                try:
                    self.browser.quit()
                except Exception as e:
                    self.logger.warning(f"Error closing browser: {e}")
    
    def _get_chrome_options(self) -> Options:
        """Configure Chrome options for optimal performance."""
        options = Options()
        
        # Performance optimizations
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        options.add_argument("--log-level=3")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-web-security")
        options.add_argument("--allow-running-insecure-content")
        options.add_argument("--headless=new")
        options.add_argument("--enable-unsafe-swiftshader")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        
        # Set download preferences
        prefs = {
            "download.default_directory": str(self.download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1 
        }
        options.add_experimental_option("prefs", prefs)
        
        return options
    
    def _wait_and_click(self, locator: tuple, timeout: int = None, use_js: bool = False) -> bool:
        """Wait for element and click it safely."""
        timeout = timeout or self.timeout
        try:
            wait = WebDriverWait(self.browser, timeout)
            element = wait.until(EC.element_to_be_clickable(locator))
            
            # Scroll element into view
            self.browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            time.sleep(0.5)  # Small delay for scroll
            
            # Try different click strategies
            if use_js:
                # Use JavaScript click to bypass interception
                self.browser.execute_script("arguments[0].click();", element)
            else:
                try:
                    # First try normal click
                    element.click()
                except Exception as e:
                    self.logger.warning(f"Normal click failed, trying JavaScript click: {e}")
                    # Fallback to JavaScript click
                    self.browser.execute_script("arguments[0].click();", element)
            
            return True
        except TimeoutException:
            self.logger.error(f"Timeout waiting for element: {locator}")
            return False
        except Exception as e:
            self.logger.error(f"Failed to click element {locator}: {e}")
            return False
    
    def _safe_send_keys(self, locator: tuple, text: str, timeout: int = None) -> bool:
        """Send keys to element safely."""
        timeout = timeout or self.timeout
        try:
            wait = WebDriverWait(self.browser, timeout)
            element = wait.until(EC.presence_of_element_located(locator))
            element.clear()
            element.send_keys(text)
            return True
        except TimeoutException:
            self.logger.error(f"Timeout waiting for input element: {locator}")
            return False
        
    def _get_target_months(self) -> List[Tuple[datetime.datetime, str]]:
        """Get the current month and previous month with formatted strings."""
        now = datetime.datetime.now()
        current_month = now
        previous_month = None
        
        if now.day <= 7:
        # Calculate previous month
            if now.month == 1:
                previous_month = now.replace(year=now.year - 1, month=12)
            else:
                previous_month = now.replace(month=now.month - 1)
        
        months = [(current_month, f"{current_month.year}年{current_month.month}月")]
        if previous_month:
            months.insert(0, (previous_month, f"{previous_month.year}年{previous_month.month}月"))
        
        return months
    
    def cleanup_old_files(self) -> bool:
        """Cleanup files matching the pattern."""
        old_files = glob.glob(self.pattern)
        
        if not old_files:
            self.logger.info("No file to clean up")
            return True
        
        success_count = 0
        for file in old_files:
            try:
                os.remove(file)
                self.logger.info(f"Cleanup old file: {os.path.basename(file)}")
                success_count += 1
            except OSError as e:
                self.logger.warning(f"Failed to delete {os.path.basename(file)}: {e}")
        
        if success_count == len(old_files):
            self.logger.info(f"Successfully cleanup {success_count} file(s)")
            return True
        else:
            self.logger.warning(f"Partially cleanup: {success_count}/{len(old_files)} file(s) deleted")
            return False

    def login(self) -> bool:
        """Handle the login process."""
        try:
            self.browser.get("https://www.einvoice.nat.gov.tw/accounts/login")
            self.recaptcha_solver = RecaptchaSolver(self.browser)
            
            # Wait for page to load completely
            WebDriverWait(self.browser, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href^="/accounts/login/b"]'))
            )
            
            # Select business login type
            if not self._wait_and_click((By.CSS_SELECTOR, 'a[href^="/accounts/login/b"]'), use_js=True):
                return False
            
            # Wait for login form to appear
            WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located((By.ID, "ban"))
            )
            
            # Handle CAPTCHA - try automated solver first, fallback to manual input
            captcha_code = None
            if self.recaptcha_solver:
                try:
                    self.logger.info("Attempting to solve CAPTCHA automatically...")
                    captcha_code = self.recaptcha_solver.solveAudioCaptcha()
                    if captcha_code:
                        self.logger.info("CAPTCHA solved automatically")
                except Exception as e:
                    self.logger.warning(f"Automatic CAPTCHA solving failed: {e}")
            
            # Fallback to manual input if automatic solving failed
            if not captcha_code:
#                captcha_code = input("Please enter the Captcha code: ").strip()
#                if not captcha_code:
#                    self.logger.error("Captcha code is required")
                    return False
            
            # Fill login form
            login_fields = [
                ((By.ID, "ban"), loginInfo.ban),
                ((By.ID, "user_id"), loginInfo.user_id),
                ((By.ID, "user_password"), loginInfo.password),
                ((By.ID, "captcha"), captcha_code)
            ]
            
            for locator, value in login_fields:
                if not self._safe_send_keys(locator, value):
                    return False
            
            if not self._wait_and_click((By.ID, 'submitBtn'), use_js=True):
                return False
            
            # Wait for login to complete and page to load
            time.sleep(5)

            if not self.browser.current_url.startswith('https://www.einvoice.nat.gov.tw/dashboard'):
                self.logger.error("Login verification failed")
                return False
            
            # Close any popup dialogs
            self._dismiss_popups()
            
            self.logger.info("Login successful")
            return True
            
        except Exception as e:
            self.logger.error(f"Login failed: {e}")
            return False
    
    def _dismiss_popups(self):
        """Try to dismiss any popup dialogs that might appear."""
        popup_selectors = [
            'button[aria-label="Close"]',
            '.modal-close',
            '.close',
            'button.btn-close',
            '[data-dismiss="modal"]'
        ]
        
        popup_texts = ['關閉', '確定', 'OK', 'Close']
        
        for selector in popup_selectors:
            if self._wait_and_click((By.CSS_SELECTOR, selector), use_js=True):
                self.logger.info(f"Closed popup with selector: {selector}")
                return True
        
        for text in popup_texts:
            if self._wait_and_click((By.XPATH, f"//button[contains(text(), '{text}')]"), use_js=True):
                self.logger.info(f"Closed popup with text: {text}")
                return True
        
        # Try pressing Escape key as last resort
        try:
            from selenium.webdriver.common.keys import Keys
            self.browser.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(0.5)
            self.logger.info("Tried to close popup with Escape key")
            return True
        except Exception:
            pass
        
        self.logger.info("No popup found to dismiss")
        return True
    
    def navigate_to_download_page(self) -> bool:
        """Navigate to the invoice download page."""
        try:
            navigation_steps = [
                (By.ID, 'headingFunctionB2B_MENU'),
                (By.ID, 'headingFunctionB2BC_SINGLE_QRY_DOWN'),
                (By.ID, 'headingFunctionBTB412W')
            ]
            
            for i, locator in enumerate(navigation_steps):
                self.logger.info(f"Clicking navigation step {i+1}: {locator[1]}")
                
                if not self._wait_and_click(locator, use_js=True):
                    self.logger.error(f"Failed to click navigation step {i+1}: {locator[1]}")
                    return False
                
                self.logger.info(f"Successfully clicked {locator[1]}")
            
            self.logger.info("Successfully navigated to download page")
            return True
            
        except Exception as e:
            self.logger.error(f"Navigation failed: {e}")
            return False
    
    def configure_search_options(self, month_date: datetime.datetime, formatted_date: str) -> bool:
        """Configure search options and filters for a specific month."""
        try:
            # Wait for page to be ready
            time.sleep(1)
            
            # Set the date input to specified month using the date picker
            try:
                # Click the date input field to open the picker
                if not self._wait_and_click((By.ID, "dp-input-date01"), use_js=True):
                    return False
                
                self.logger.info("Date picker opened")
                time.sleep(0.5)
                
                # Wait for the overlay to appear
                WebDriverWait(self.browser, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".dp__overlay.dp--overlay-relative"))
                )
                
                # Check if we need to change the year first
                current_year_button = self.browser.find_element(By.CSS_SELECTOR, ".dp__btn.dp--year-select")
                current_year_text = current_year_button.text.strip()
                target_year = f"{month_date.year}年"
                
                if current_year_text != target_year:
                    self.logger.info(f"Need to change year from {current_year_text} to {target_year}")
                    
                    if not self._wait_and_click((By.CSS_SELECTOR, ".dp__btn.dp--year-select"), use_js=True):
                        return False
                    
                    time.sleep(0.5)
                    
                    # Navigate to the correct year using arrow buttons
                    current_year_num = int(current_year_text.replace('年', ''))
                    target_year_num = month_date.year
                    
                    if target_year_num < current_year_num:
                        # Click previous year button
                        for _ in range(current_year_num - target_year_num):
                            if not self._wait_and_click((By.CSS_SELECTOR, ".dp__btn.dp--arrow-btn-nav[aria-label='Previous year']"), use_js=True):
                                return False
                            time.sleep(0.3)
                    elif target_year_num > current_year_num:
                        # Click next year button
                        for _ in range(target_year_num - current_year_num):
                            if not self._wait_and_click((By.CSS_SELECTOR, ".dp__btn.dp--arrow-btn-nav[aria-label='Next year']"), use_js=True):
                                return False
                            time.sleep(0.3)
                
                # Now select the correct month
                target_month_text = f"{month_date.month}月"
                
                if not self._wait_and_click((By.CSS_SELECTOR, f'div[data-test="{target_month_text}"]'), use_js=True):
                    return False
                
                self.logger.info(f"Selected month: {target_month_text}")
                time.sleep(0.5)
                
                # Verify the selection was successful by checking the input value
                try:
                    date_input = self.browser.find_element(By.ID, "dp-input-date01")
                    updated_value = date_input.get_attribute("value")
                    self.logger.info(f"Date input updated to: {updated_value}")
                except Exception as e:
                    self.logger.warning(f"Could not verify date input value: {e}")
                
            except Exception as e:
                self.logger.error(f"Failed to set date using date picker: {e}")
                # Try to close any open overlay
                try:
                    self.browser.execute_script("document.body.click();")
                except Exception:
                    pass
                return False
            
            radio_options = [
                (By.ID, "queryInvType_1"),
                (By.ID, "businessType_1")
            ]
            
            for locator in radio_options:
                if not self._wait_and_click(locator, use_js=True):
                    self.logger.error(f"Failed to select radio button {locator[1]}")
                    return False
                self.logger.info(f"Selected radio button: {locator[1]}")
            
            if not self._wait_and_click((By.CSS_SELECTOR, 'button[title="查詢"]'), use_js=True):
                self.logger.error("Failed to click search button")
                return False
            
            self.logger.info("Search button clicked")
            
            # Wait for search results to load
            try:
                WebDriverWait(self.browser, 2.5).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//span[text()='查詢成功。']")
                    )
                )

            except TimeoutException:
                self.logger.info(
                    f"No result for {month_date.year}年 {month_date.month}月"
                )
                return False
            
            # Set page size to maximum
            try:
                select_element = WebDriverWait(self.browser, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'select[title="分頁"]'))
                )
                self.browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_element)
                time.sleep(0.5)
                select = Select(select_element)
                select.select_by_value("1000")
                self.logger.info("Page size set to 1000")
                time.sleep(1)  # Wait for page to reload with new size
            except TimeoutException:
                self.logger.warning("Could not find page size selector - continuing without changing page size")
            except Exception as e:
                self.logger.warning(f"Could not set page size: {e}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to configure search options: {e}")
            return False

    def download_invoices(self) -> bool:
        """Select all invoices and download from all pages."""
        try:
            page_count = 0
            
            while True:
                page_count += 1
                self.logger.info(f"Processing page {page_count}")
                
                # Scroll to top first to ensure we can see the elements
                self.browser.execute_script("window.scrollTo(0, 0);")
                time.sleep(0.5)
                
                try:
                    # Check if checkbox is already selected first
                    checkbox = WebDriverWait(self.browser, 15).until(
                        EC.presence_of_element_located((By.ID, "checkbox-all"))
                    )
                    
                    if not checkbox.is_selected():
                        if not self._wait_and_click((By.ID, "checkbox-all"), use_js=True):
                            self.logger.error(f"Failed to click select all checkbox on page {page_count}")
                            return False
                        self.logger.info(f"Select all checkbox clicked on page {page_count}")
                    else:
                        self.logger.info(f"Select all checkbox already selected on page {page_count}")
                    
                    time.sleep(0.5)
                    
                except Exception as e:
                    self.logger.error(f"Failed to handle checkbox on page {page_count}: {e}")
                    return False
                
                if not self._wait_and_click((By.CSS_SELECTOR, 'button[title="下載Excel檔"]'), use_js=True):
                    self.logger.error(f"Failed to click download button on page {page_count}")
                    return False
                
                self.logger.info(f"Download button clicked successfully on page {page_count}")
                self.total_downloads += 1

                # Wait for Download to complete
                time.sleep(5)
                
                # Check if there's a next page
                try:
                    next_button = WebDriverWait(self.browser, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'button[title="下一頁"]'))
                    )
                    
                    # Check if next button is disabled (no more pages)
                    if next_button.get_attribute('disabled'):
                        self.logger.info(f"No more pages. Completed processing {page_count} page(s)")
                        break
                    else:
                        if not self._wait_and_click((By.CSS_SELECTOR, 'button[title="下一頁"]'), use_js=True):
                            self.logger.error(f"Failed to click next page button on page {page_count}")
                            break
                        
                        self.logger.info(f"Moving to page {page_count + 1}")
                        time.sleep(2)  # Wait for page to load
                        
                except Exception as e:
                    self.logger.error(f"Failed to check next page button on page {page_count}: {e}")
                    # If we can't find the next button, assume we're done
                    break
            
            return True
            
        except Exception as e:
            self.logger.error(f"Download failed: {e}")
            return False
    
    def wait_for_download(self, max_wait_time: int = 60) -> str:
        """Wait for download to complete and return file path."""
        
        start_time = time.time()
        while time.time() - start_time < max_wait_time:
            matched_files = glob.glob(self.pattern)
            if len(matched_files) == self.total_downloads:
                for file in matched_files:
                    self.logger.info(f"Found the matching file: {os.path.basename(file)}")
                return matched_files
            time.sleep(0.5)
        
        self.logger.error("Download timeout - files not found")
        return None
    
    def run(self) -> str:
        """Execute the complete download process."""
        self.logger.info("Starting invoice download process")
        
        with self.get_browser():
            if not self.login():
                raise Exception("Login failed")
            
            if not self.navigate_to_download_page():
                raise Exception("Navigation failed")
            
            # Get target months (previous and current)
            target_months = self._get_target_months()

            if not self.cleanup_old_files():
                raise Exception("Cleanup failed")

            for i, (month_date, formatted_date) in enumerate(target_months):
                self.logger.info(f"Processing month {i+1}: {formatted_date}")
                
                # For subsequent months, we need to refresh the page or reset the form
                if i > 0:
                    self.logger.info("Refreshing page for next month...")
                    self.browser.refresh()
                    time.sleep(2)
            
                # Configure search for this specific month
                if not self.configure_search_options(month_date, formatted_date):
                    self.logger.error(f"Search configuration failed for {formatted_date}")
                    continue
                
                # Download invoices for this month
                if not self.download_invoices():
                    self.logger.error(f"Download initiation failed for {formatted_date}")
                    continue

            downloaded_file = self.wait_for_download()
            if not downloaded_file:
                raise Exception("Download did not complete.")
            
            return downloaded_file


def main():
    """Main execution function."""
    try:
        downloader = InvoiceDownloader()
        downloader.run()
        downloader.logger.info("EInvoice Downloader Success!\n")
        
    except Exception as e:
        downloader.logger.error(f"Error: {e}\n")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())