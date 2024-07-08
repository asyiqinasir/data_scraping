import os
import requests
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def download_pdf(url, save_path):
    """Download a PDF file from the given URL and save it to the specified path."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        with open(save_path, 'wb') as f:
            f.write(response.content)
        logging.info(f"Downloaded PDF from {url} to {save_path}")
    except requests.RequestException as e:
        logging.error(f"Failed to download PDF from {url}: {e}")
        raise

def setup_webdriver():
    """Setup and return the Chrome WebDriver using webdriver_manager."""
    try:
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        logging.info("Chrome WebDriver initialized successfully")
        return driver
    except Exception as e:
        logging.error(f"Failed to initialize Chrome WebDriver: {e}")
        raise

def get_pdf_urls(driver, url):
    """Retrieve PDF URLs from the specified webpage."""
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 30)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        pdf_links = driver.find_elements(By.XPATH, '//a[contains(@href, ".pdf")]')
        pdf_urls = [link.get_attribute('href') for link in pdf_links]
        logging.info(f"Found {len(pdf_urls)} PDF URLs on the page")
        return pdf_urls
    except Exception as e:
        logging.error(f"Failed to retrieve PDF URLs from {url}: {e}")
        raise

def main():
    """Main function to download PDFs."""
    url = 'https://www.mof.gov.my/portal/arkib/perbelanjaan/exp2013.html'
    pdf_dir = 'input/2013'

    os.makedirs(pdf_dir, exist_ok=True)

    driver = setup_webdriver()

    try:
        pdf_urls = get_pdf_urls(driver, url)
        
        for pdf_url in pdf_urls:
            pdf_name = pdf_url.split('/')[-1]
            pdf_path = os.path.join(pdf_dir, pdf_name)

            if not os.path.exists(pdf_path):
                download_pdf(pdf_url, pdf_path)
    
    except Exception as e:
        logging.error(f"An error occurred: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
