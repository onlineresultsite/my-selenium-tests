import pytest
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService

@pytest.mark.test
def test_google_search():
    chrome_driver_path = '/usr/local/bin/chromedriver'  # Update this with your actual ChromeDriver path

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # Optional: Run Chrome in headless mode
    options.add_argument('--no-sandbox')  # Required for running as root user

    service = ChromeService(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=options)

    # Replace with your test logic
    driver.get("https://www.google.com")
    assert "Google" in driver.title

    driver.quit()
