import pytest
import openpyxl
from openpyxl.styles import Font, PatternFill
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Global variables for Excel handling
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Test Results"

# Row counter for Excel sheet
row_counter = 1

# URL, username, and password
URL = "http://13.53.206.233:8000/login/"  # Replace with your login URL
USERNAME = "newtest"
PASSWORD = "akTR@300"

# Login function
def login(driver, username, password):
    driver.get(URL)

    try:
        # Wait for username input field to be visible
        username_input = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, "//input[@name='username']"))
        )
        username_input.send_keys(username)

        # Find password input field
        password_input = driver.find_element(By.XPATH, "//input[@name='password']")
        password_input.send_keys(password)

        # Click on the login button
        login_button = driver.find_element(By.XPATH, "//button[@type='submit' and contains(@class, 'btn-primary')]")
        login_button.click()

        # Wait for the welcome message or dashboard to appear after successful login
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, "//h1[contains(text(), 'Welcome')]"))
            # Adjust the XPath according to the welcome message or element you expect
        )

        # Write results to Excel file
        global row_counter
        row_counter += 1
        sheet[f"A{row_counter}"] = "Dashboard"
        sheet[f"B{row_counter}"] = "Success"

        # Save the Excel file
        wb.save("test_results.xlsx")

        # Print a message
        print("Test results saved to test_results.xlsx")

    except TimeoutException as e:
        print(f"TimeoutException occurred: {e}")
        assert False, "Timeout waiting for element."

    finally:
        # Add any necessary cleanup or verification steps here
        pass

# Pytest fixtures for setup and teardown
@pytest.fixture(scope="module")
def setup():
    # Setup the WebDriver
    driver = webdriver.Chrome()  # Use the browser you prefer
    yield driver
    # Teardown
    driver.quit()

# Test function with parameterization
@pytest.mark.parametrize("username, password", [
    (USERNAME, PASSWORD),  # Use the correct credentials here
])
def test_login(setup, username, password):
    driver = setup
    login(driver, username, password)

# Run the tests if this script is executed directly
if __name__ == "__main__":
    row_counter = 1  # Initialize the row_counter
    pytest.main()
