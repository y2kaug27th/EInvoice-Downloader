import os
import asyncio
import aiohttp
import datetime
import subprocess
import tempfile
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import whisper

class RecaptchaSolver:
    def __init__(self, driver, model_size="base"):
        self.driver = driver
        # Initialize Whisper model
        # Available models: tiny, base, small, medium, large
        # "base" is a good balance between speed and accuracy
        print(f"Loading Whisper model: {model_size}")
        self.model = whisper.load_model(model_size)
        print("Whisper model loaded successfully")

    async def download_audio(self, url, path):
        async with aiohttp.ClientSession() as session:
            async with session.get(url) as response:
                with open(path, 'wb') as f:
                    f.write(await response.read())
        print("Downloaded audio asynchronously.")

    def recognize_audio_with_whisper(self, audio_path):
        """Recognize audio using OpenAI Whisper"""
        try:
            # Whisper can handle various audio formats directly
            print(f"Transcribing audio: {audio_path}")
            
            # Transcribe with language hint for better accuracy
            # Set language to Chinese for better recognition of Chinese numbers
            result = self.model.transcribe(
                audio_path,
                language="zh",  # Chinese language hint
                task="transcribe"
            )
            
            recognized_text = result["text"].strip()
            print(f"Whisper recognition result: {recognized_text}")
            
            return recognized_text
            
        except Exception as e:
            print(f"Whisper recognition failed: {e}")
            return None

    def convert_audio_format(self, input_path, output_path):
        """Convert audio to a format that works well with Whisper"""
        command = [
            "ffmpeg",
            "-y",  # overwrite if file exists
            "-i", input_path,
            "-ar", "16000",  # 16kHz sample rate (Whisper's preferred)
            "-ac", "1",      # Mono
            "-c:a", "pcm_s16le",  # 16-bit PCM
            output_path
        ]
        try:
            subprocess.run(command, capture_output=True, text=True, check=True)
            print("Audio conversion successful")
            return True
        except subprocess.CalledProcessError as e:
            print("FFmpeg conversion failed.")
            print("STDOUT:", e.stdout)
            print("STDERR:", e.stderr)
            return False

    def solveAudioCaptcha(self):
        try:
            self.driver.switch_to.default_content()

            # Click on the audio button
            audio_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[title="語音播放圖形驗證碼"]'))
            )
            audio_button.click()

            # Get the audio source URL
            audio_source = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'audio'))
            ).get_attribute('src')
            print("Audio source URL detected")
            
            # Create temporary files
            with tempfile.TemporaryDirectory() as temp_dir:
                timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
                path_to_original = os.path.join(temp_dir, f"{timestamp}.mp3")
                path_to_converted = os.path.join(temp_dir, f"{timestamp}_converted.wav")
                
                # Download the audio asynchronously
                asyncio.run(self.download_audio(audio_source, path_to_original))

                # Convert audio format for better compatibility
                if not self.convert_audio_format(path_to_original, path_to_converted):
                    print("Audio conversion failed, trying with original file")
                    audio_file_to_use = path_to_original
                else:
                    audio_file_to_use = path_to_converted

                # Recognize the audio using Whisper
                captcha_text = None
                for attempt in range(3):
                    try:
                        recognized_text = self.recognize_audio_with_whisper(audio_file_to_use)
                        if recognized_text:
                            # Clean up the text (remove spaces, punctuation)
                            captcha_text = ''.join(char for char in recognized_text if char.isalnum() or char in '一二三四五六七八九零')
                            print(f"Cleaned recognized text: {captcha_text}")
                            break
                    except Exception as e:
                        print(f"Whisper recognition attempt {attempt + 1} failed: {e}")

            if not captcha_text:
                print("Failed to solve audio CAPTCHA with Whisper")
                return None
            
            # Convert Chinese numbers to digits
            captcha_number = self.convert_chinese_to_digits(captcha_text)
            
            if not captcha_number:
                print("Failed to convert recognized text to numbers")
                return None

            print(f"Final CAPTCHA number: {captcha_number}")
            return captcha_number

        except Exception as e:
            print(f"An error occurred while solving audio CAPTCHA: {e}")
            self.driver.switch_to.default_content()
            raise

        finally:
            # Always switch back to the main content
            self.driver.switch_to.default_content()

    def convert_chinese_to_digits(self, text):
        """Convert Chinese numbers to digits"""
        captcha_number = ''
        
        # Handle both traditional and simplified Chinese numbers
        chinese_to_digit = {
            '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
            '六': '6', '七': '7', '八': '8', '九': '9', '零': '0',
            # Also handle potential variations
            '1': '1', '2': '2', '3': '3', '4': '4', '5': '5',
            '6': '6', '7': '7', '8': '8', '9': '9', '0': '0',
            # Other circumstances
            'E': '1'
        }
        
        for char in text:
            if char in chinese_to_digit:
                captcha_number += chinese_to_digit[char]
            else:
                # If we encounter an unrecognized character, log it but continue
                print(f"Unrecognized character: {char}")
        
        # Validate the result (typical CAPTCHA length)
        if len(captcha_number) == 5:
            return captcha_number
        else:
            print(f"Invalid CAPTCHA length: {len(captcha_number)}")
            return None