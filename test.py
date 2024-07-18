import pytest
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import time

@pytest.mark.test
def test_google_search():
    chrome_driver_path = '/path/to/your/chromedriver'  # Update this with your actual ChromeDriver path

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # Optional: Run Chrome in headless mode
    options.add_argument('--no-sandbox')  # Required for running as root user

    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get("https://www.google.com/")
        search_box = driver.find_element_by_name("q")
        search_box.send_keys("OpenAI ChatGPT")
        search_box.send_keys(Keys.RETURN)
        time.sleep(5)
        assert "OpenAI ChatGPT" in driver.title
    finally:
        driver.quit()
