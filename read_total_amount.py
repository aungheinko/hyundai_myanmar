import os
import time
import requests  # Import requests to download images
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Function to download images
def download_images_from_facebook(page_url, download_folder):
    # Set up Chrome options
    chrome_options = Options()
    # chrome_options.add_argument("--headless")  # Uncomment this line to run in headless mode

    # Create a new instance of the Chrome driver
    service = Service(r'C:\Users\asservices012\Downloads\test\chromedriver-win64\chromedriver.exe')  # Replace with the path to your ChromeDriver
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        print(f"Visiting {page_url}...")
        driver.get(page_url)
        time.sleep(5)  # Wait for the page to load

        # Scroll to the bottom of the page to load more images
        scroll_pause_time = 2
        last_height = driver.execute_script("return document.body.scrollHeight")

        while True:
            # Scroll down to the bottom
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(scroll_pause_time)

            # Calculate new scroll height and compare with last height
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # Find all image elements
        images = driver.find_elements(By.TAG_NAME, 'img')
        print(f"Found {len(images)} images.")

        # Create folder if it doesn't exist
        if not os.path.exists(download_folder):
            os.makedirs(download_folder)

        # Download images
        for index, img in enumerate(images):
            # Get the URL for the image, checking various attributes for higher quality
            img_url = img.get_attribute('src') or img.get_attribute('data-src') or img.get_attribute('srcset')
            if img_url:
                try:
                    # Download the image
                    img_data = requests.get(img_url).content
                    with open(os.path.join(download_folder, f'image_{index + 1}.jpg'), 'wb') as handler:
                        handler.write(img_data)
                    print(f"Downloaded: image_{index + 1}.jpg")
                    time.sleep(1)  # Pause to avoid rate limiting
                except Exception as e:
                    print(f"Failed to download {img_url}: {e}")
    finally:
        driver.quit()

# Main execution
if __name__ == "__main__":
    try:
        page_url = "https://www.facebook.com/myanmarbeautiful"  # Replace with your Facebook page URL
        download_folder = "facebook_images"  # Folder to save images
        download_images_from_facebook(page_url, download_folder)
    except Exception as e:
        print(f"An error occurred: {e}")
