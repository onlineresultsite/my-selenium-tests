import pytest
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

@pytest.mark.test
def test_google_search():
    driver = webdriver.Chrome()
    try:
        driver.get("https://www.google.com/")
        search_box = driver.find_element_by_name("q")
        search_box.send_keys("OpenAI ChatGPT")
        search_box.send_keys(Keys.RETURN)
        time.sleep(5)
        assert "OpenAI ChatGPT" in driver.title
    finally:
        driver.quit()
