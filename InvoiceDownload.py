import glob
import time
import datetime
import logging
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
                 timeout: int = 30,
                 recaptcha_solver=None):
        self.webdriver_path = webdriver_path
        self.download_dir = Path(download_dir)
        self.timeout = timeout
        self.browser = None
        self.total_downloads = 0
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
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-web-security")
        options.add_argument("--allow-running-insecure-content")
        
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
            
            # Submit login
            if not self._wait_and_click((By.ID, 'submitBtn'), use_js=True):
                return False
            
            # Wait for login to complete and page to load
            time.sleep(5)

            if not self.browser.current_url.startswith('https://www.einvoice.nat.gov.tw/dashboard'):
                self.logger.error("Login verification failed.")
                return False
            
            # Close any popup dialogs (try multiple strategies)
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
        
        # Try CSS selectors first
        for selector in popup_selectors:
            try:
                element = WebDriverWait(self.browser, 3).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
                self.browser.execute_script("arguments[0].click();", element)
                time.sleep(0.5)
                self.logger.info(f"Closed popup with selector: {selector}")
                return True
            except Exception:
                continue
        
        # Try text-based buttons
        for text in popup_texts:
            try:
                element = WebDriverWait(self.browser, 3).until(
                    EC.element_to_be_clickable((By.XPATH, f"//button[contains(text(), '{text}')]"))
                )
                self.browser.execute_script("arguments[0].click();", element)
                time.sleep(0.5)
                self.logger.info(f"Closed popup with text: {text}")
                return True
            except Exception:
                continue
        
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
            # Wait for any overlays or loading to complete
            
            navigation_steps = [
                (By.ID, 'headingFunctionB2B_MENU'),
                (By.ID, 'headingFunctionB2BC_SINGLE_QRY_DOWN'),
                (By.ID, 'headingFunctionBTB412W')
            ]
            
            for i, locator in enumerate(navigation_steps):
                self.logger.info(f"Clicking navigation step {i+1}: {locator[1]}")
                
                # Wait for element to be present and visible
                wait = WebDriverWait(self.browser, 15)
                element = wait.until(EC.presence_of_element_located(locator))
                
                # Try multiple click strategies
                clicked = False
                
                # Strategy 1: JavaScript click (most reliable for intercepted elements)
                try:
                    self.browser.execute_script("arguments[0].click();", element)
                    clicked = True
                    self.logger.info(f"Successfully clicked {locator[1]} with JavaScript")
                except Exception as e:
                    self.logger.warning(f"JavaScript click failed for {locator[1]}: {e}")
                
                # Strategy 2: Action chains if JS failed
                if not clicked:
                    try:
                        from selenium.webdriver.common.action_chains import ActionChains
                        actions = ActionChains(self.browser)
                        actions.move_to_element(element).click().perform()
                        clicked = True
                        self.logger.info(f"Successfully clicked {locator[1]} with ActionChains")
                    except Exception as e:
                        self.logger.warning(f"ActionChains click failed for {locator[1]}: {e}")
                
                # Strategy 3: Regular click as last resort
                if not clicked:
                    try:
                        element.click()
                        clicked = True
                        self.logger.info(f"Successfully clicked {locator[1]} with regular click")
                    except Exception as e:
                        self.logger.error(f"All click strategies failed for {locator[1]}: {e}")
                        return False
                
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
                # Find and click the date input field to open the picker
                date_input = WebDriverWait(self.browser, 10).until(
                    EC.element_to_be_clickable((By.ID, "dp-input-date01"))
                )
                
                self.browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", date_input)
                time.sleep(0.5)
                
                # Click to open the date picker overlay
                self.browser.execute_script("arguments[0].click();", date_input)
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
                    
                    # Click on year to enter year selection mode
                    self.browser.execute_script("arguments[0].click();", current_year_button)
                    time.sleep(0.5)
                    
                    # Navigate to the correct year using arrow buttons
                    current_year_num = int(current_year_text.replace('年', ''))
                    target_year_num = month_date.year
                    
                    if target_year_num < current_year_num:
                        # Click previous year button
                        for _ in range(current_year_num - target_year_num):
                            prev_year_btn = self.browser.find_element(By.CSS_SELECTOR, ".dp__btn.dp--arrow-btn-nav[aria-label='Previous year']")
                            self.browser.execute_script("arguments[0].click();", prev_year_btn)
                            time.sleep(0.3)
                    elif target_year_num > current_year_num:
                        # Click next year button
                        for _ in range(target_year_num - current_year_num):
                            next_year_btn = self.browser.find_element(By.CSS_SELECTOR, ".dp__btn.dp--arrow-btn-nav[aria-label='Next year']")
                            self.browser.execute_script("arguments[0].click();", next_year_btn)
                            time.sleep(0.3)
                
                # Now select the correct month
                target_month_text = f"{month_date.month}月"
                
                # Find and click the target month cell
                month_cell = WebDriverWait(self.browser, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, f'div[data-test="{target_month_text}"]'))
                )
                
                self.browser.execute_script("arguments[0].click();", month_cell)
                self.logger.info(f"Selected month: {target_month_text}")
                time.sleep(0.5)
                
                # The overlay should close automatically after selection
                # Wait a moment to ensure the selection is registered
                time.sleep(0.5)
                
                # Verify the selection was successful by checking the input value
                try:
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
            
            # Select radio buttons
            radio_options = [
                (By.ID, "queryInvType_1"),
                (By.ID, "businessType_1")
            ]
            
            for locator in radio_options:
                try:
                    wait = WebDriverWait(self.browser, 10)
                    element = wait.until(EC.element_to_be_clickable(locator))
                    self.browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                    time.sleep(0.5)
                    self.browser.execute_script("arguments[0].click();", element)
                    self.logger.info(f"Selected radio button: {locator[1]}")
                except Exception as e:
                    self.logger.error(f"Failed to select radio button {locator[1]}: {e}")
                    return False
            
            # Click search button
            try:
                search_button = WebDriverWait(self.browser, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[title="查詢"]'))
                )
                self.browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", search_button)
                time.sleep(0.5)
                self.browser.execute_script("arguments[0].click();", search_button)
                self.logger.info("Search button clicked")
            except Exception as e:
                self.logger.error(f"Failed to click search button: {e}")
                return False
            
            # Wait for search results to load
            time.sleep(2.5)
            
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
                
                # Find and select all checkboxes on current page
                try:
                    checkbox = WebDriverWait(self.browser, 15).until(
                        EC.presence_of_element_located((By.ID, "checkbox-all"))
                    )
                    
                    # Scroll the checkbox into view and center it
                    self.browser.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", checkbox)
                    time.sleep(0.5)
                    
                    # Check if it's already selected
                    if not checkbox.is_selected():
                        self.browser.execute_script("arguments[0].click();", checkbox)
                        self.logger.info(f"Select all checkbox clicked on page {page_count}")
                    else:
                        self.logger.info(f"Select all checkbox already selected on page {page_count}")
                    
                    time.sleep(0.5)
                    
                except Exception as e:
                    self.logger.error(f"Failed to select checkbox on page {page_count}: {e}")
                    return False
                
                # Find and click download button for current page
                try:
                    download_button = WebDriverWait(self.browser, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'button[title="下載Excel檔"]'))
                    )
                    
                    # Scroll the download button into view and center it
                    self.browser.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", download_button)
                    time.sleep(0.5)
                    
                    # Make sure button is clickable
                    WebDriverWait(self.browser, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[title="下載Excel檔"]'))
                    )
                    
                    # Click using JavaScript to avoid interception
                    self.browser.execute_script("arguments[0].click();", download_button)
                    self.logger.info(f"Download button clicked successfully on page {page_count}")
                    self.total_downloads += 1

                    # Wait for Download to complete
                    time.sleep(5)
                    
                except Exception as e:
                    self.logger.error(f"Failed to click download button on page {page_count}: {e}")
                    return False
                
                # Check if there's a next page
                try:
                    next_button = WebDriverWait(self.browser, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'button[title="下一頁"]'))
                    )
                    
                    # Check if next button is disabled (no more pages)
                    if next_button.get_attribute('disabled'):
                        self.logger.info(f"No more pages. Completed processing {page_count} pages")
                        break
                    else:
                        # Click next page button
                        self.browser.execute_script("arguments[0].click();", next_button)
                        self.logger.info(f"Moving to page {page_count + 1}")
                        time.sleep(2)  # Wait for page to load
                        
                except Exception as e:
                    self.logger.error(f"Failed to check next page button on page {page_count}: {e}")
                    # If we can't find the next button, assume we're done
                    break
            
            self.logger.info(f"Download process completed successfully. Total pages processed: {page_count}, Total downloads: {self.total_downloads}")
            return True
            
        except Exception as e:
            self.logger.error(f"Download failed: {e}")
            return False
    
    def wait_for_download(self, max_wait_time: int = 60) -> str:
        """Wait for download to complete and return file path."""
        prefix = f"14001199_IN_{datetime.datetime.today().strftime('%Y%m%d%H')}"
        pattern = str(self.download_dir / f"{prefix}*.xls")
        
        start_time = time.time()
        while time.time() - start_time < max_wait_time:
            matched_files = glob.glob(pattern)
            if len(matched_files) >= self.total_downloads:
                download_files = matched_files[-self.total_downloads:]
                for file in download_files:
                    self.logger.info(f"Found the matching files: {file}")
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

            for i, (month_date, formatted_date) in enumerate(target_months):
                self.logger.info(f"Processing month {i+1}/2: {formatted_date}")
                
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
        downloader.logger.info("EInvoice Downloader Success!")
        
    except Exception as e:
        downloader.logger.error(f"Error: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())