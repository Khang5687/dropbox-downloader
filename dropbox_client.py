from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
import time
from pathlib import Path
import os
import argparse
import sys
from tqdm import tqdm

def download_first_file(url, output_dir, debug=False, use_alt_method=False, user_data_dir="/tmp/chrome-debug", progress_bar=None, file_label=""):
    """
    Download the first file from a Dropbox shared folder.
    
    Args:
        url: Dropbox shared folder URL
        output_dir: Directory to save the downloaded file
        debug: If True, print verbose debug messages
        use_alt_method: If True, use button-click method instead of URL-based download
        user_data_dir: Chrome user data directory for session persistence
        progress_bar: Optional tqdm progress bar instance
        file_label: Label for the file being downloaded (e.g., UPC)
        
    Returns:
        Path to downloaded file or None if failed
    """
    def log(msg):
        if debug:
            if progress_bar:
                progress_bar.write(msg)
            else:
                print(msg)
    
    def update_progress(msg):
        """Update progress bar description"""
        if progress_bar:
            progress_bar.set_description(f"{file_label}: {msg}")
    
    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    # Load user agent from file
    user_agent_file = Path("useragent.txt")
    if user_agent_file.exists():
        user_agent = user_agent_file.read_text().strip()
        chrome_options.add_argument(f"--user-agent={user_agent}")
    else:
        log("Warning: useragent.txt not found, using default user agent")
    
    # Headless mode with Windows compatibility fixes
    chrome_options.add_argument("--headless=new")  # Use new headless mode (more stable)
    chrome_options.add_argument("--no-sandbox")  # Required for Windows in some environments
    chrome_options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
    chrome_options.add_argument("--disable-gpu")  # Disable GPU hardware acceleration
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--window-size=1920,1080")  # Set window size for headless
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-dev-tools")
    chrome_options.add_argument("--remote-debugging-port=0")  # Use random port to avoid conflicts
    
    # Use unique user data directory to prevent conflicts between threads
    chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
    
    # Configure download directory and preferences
    output_path = Path(output_dir).resolve()
    output_path.mkdir(parents=True, exist_ok=True)
    
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": str(output_path),
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True,
    })
    
    # Suppress Selenium logs
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    # Set up Chrome service with explicit log configuration
    service = Service()
    service.log_path = os.devnull  # Suppress ChromeDriver logs
    
    driver = None
    max_retries = 3
    for attempt in range(max_retries):
        try:
            driver = webdriver.Chrome(service=service, options=chrome_options)
            break
        except Exception as e:
            if attempt < max_retries - 1:
                log(f"Chrome launch attempt {attempt + 1} failed, retrying...")
                time.sleep(2)  # Wait before retry
            else:
                log(f"Failed to launch Chrome after {max_retries} attempts: {e}")
                raise

    try:
        update_progress("Loading page")
        log(f"Navigating to: {url}")
        driver.get(url)

        update_progress("Waiting for content")
        log("Waiting for Dropbox grid to load...")
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="sl-grid-body"]'))
        )
        
        # Handle cookie consent banner if present
        try:
            log("Checking for cookie consent banner...")
            consent_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="accept_all_cookies_button"]'))
            )
            consent_btn.click()
            log("Cookie consent accepted")
            time.sleep(1)
        except TimeoutException:
            log("No cookie banner detected")
        
        # Locate the first file card
        update_progress("Locating file")
        log("Locating the first file in the grid...")
        grid = driver.find_element(By.CSS_SELECTOR, '[data-testid="sl-grid-body"]')
        
        # Wait for at least one card to appear
        WebDriverWait(driver, 10).until(
            lambda d: len(grid.find_elements(By.CSS_SELECTOR, 'li._sl-card_to1nz_25')) > 0
        )
        
        first_card = grid.find_elements(By.CSS_SELECTOR, 'li._sl-card_to1nz_25')[0]
        
        # Get file name and link for logging
        try:
            file_link = first_card.find_element(By.CSS_SELECTOR, '[data-testid="grid-link"]')
            file_name = file_link.text
            preview_url = file_link.get_attribute('href')
            log(f"First file found: {file_name}")
        except:
            file_name = "unknown"
            preview_url = None
            log("First file found (name could not be retrieved)")
        
        # Download using URL-based method or button click
        update_progress("Starting download")
        if use_alt_method:
            # Button-click method (alternative)
            log("Using button-click download method...")
            
            # Hover over the card to reveal download controls
            log("Hovering over file card to reveal download button...")
            ActionChains(driver).move_to_element(first_card).pause(0.5).perform()
            
            # Wait for and click the download button
            log("Clicking download button...")
            try:
                download_btn = WebDriverWait(first_card, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, './/button[.//svg[@aria-label="Download"]]')
                    )
                )
                download_btn.click()
            except ElementClickInterceptedException:
                # Fallback to JavaScript click if regular click is intercepted
                log("Regular click intercepted, using JavaScript click...")
                download_btn = first_card.find_element(By.XPATH, './/button[.//svg[@aria-label="Download"]]')
                driver.execute_script("arguments[0].click();", download_btn)
            
            log(f"Download initiated for: {file_name}")
        else:
            # URL-based download: replace dl=0 with dl=1 (default)
            log("Using URL-based download method...")
            download_url = preview_url.replace('dl=0', 'dl=1')
            log(f"Navigating to download URL: {download_url}")
            driver.get(download_url)
            log(f"Download initiated for: {file_name}")
        
        # Wait for download to complete
        update_progress("Downloading")
        log("Waiting for download to complete...")
        timeout = 120
        start_time = time.time()
        downloaded_file = None
        
        # Get initial file list
        initial_files = set(f.name for f in output_path.iterdir() if f.is_file())
        
        while time.time() - start_time < timeout:
            elapsed = int(time.time() - start_time)
            
            # Check if any .crdownload files exist (Chrome's in-progress download extension)
            crdownload_files = list(output_path.glob("*.crdownload"))
            if crdownload_files:
                # Try to get file size for progress indication
                try:
                    size_mb = crdownload_files[0].stat().st_size / (1024 * 1024)
                    update_progress(f"Downloading ({size_mb:.1f} MB, {elapsed}s)")
                except:
                    update_progress(f"Downloading ({elapsed}s)")
                log("Download in progress...")
                time.sleep(1)
                continue
            
            # Check if any new files were downloaded
            current_files = set(f.name for f in output_path.iterdir() if f.is_file())
            new_files = current_files - initial_files
            
            if new_files:
                downloaded_file = output_path / list(new_files)[0]
                update_progress("Complete")
                log(f"Download complete: {downloaded_file}")
                break
            
            time.sleep(1)
        else:
            update_progress("Timeout")
            log("Download timeout - file may still be downloading")
            return None
        
        return downloaded_file

    except Exception as e:
        log(f"Error during download: {str(e)}")
        return None
    finally:
        driver.quit()


def main():
    """CLI entry point for testing download_dropbox.py standalone"""
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Download first file from Dropbox shared folder')
    parser.add_argument('url', help='Dropbox shared folder URL')
    parser.add_argument('--alt', action='store_true', 
                       help='Use button-click download method instead of URL-based download (default)')
    parser.add_argument('--debug', action='store_true',
                       help='Enable verbose debug output')
    parser.add_argument('--output', default='downloads',
                       help='Output directory for downloaded files (default: downloads)')
    args = parser.parse_args()
    
    result = download_first_file(
        url=args.url,
        output_dir=args.output,
        debug=args.debug,
        use_alt_method=args.alt,
        user_data_dir="/tmp/chrome-debug2"
    )
    
    if result:
        print(f"✓ Downloaded: {result}")
    else:
        print("✗ Download failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
