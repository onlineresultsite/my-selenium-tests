import pytest
import openpyxl
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys

# Initialize the workbook and sheet outside the functions
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Test Results"

row_counter = 1

@pytest.fixture(scope="module")
def setup():
    # Setup the WebDriver
    driver = webdriver.Chrome()  # Use the browser you prefer
    yield driver
    # Teardown
    driver.quit()

@pytest.fixture
def login_setup():
    # Placeholder for any additional setup actions specific to each user login
    pass

def login(driver, username, password, otp1, otp2, otp3, otp4, role):
    driver.get("https://uat-admin.kaabilfinance.com/employee/login")

    try:
        # Wait for username input field to be visible
        username_input = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, "//input[@name='empid']"))
        )
        username_input.send_keys(username)

        # Find password input field
        password_input = driver.find_element(By.XPATH, "//input[@name='password']")
        password_input.send_keys(password)

        # Click on the login button
        login_button = driver.find_element(By.XPATH, "//button[text()='Log In']")
        login_button.click()

        # Wait for the OTP input fields to become visible and enter OTPs
        otp_input1 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div/div/div/div/div/div/div[2]/div/form/div/div[1]/div/input[1]"))
        )
        otp_input1.send_keys(otp1)

        otp_input2 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div/div/div/div/div/div/div[2]/div/form/div/div[1]/div/input[2]"))
        )
        otp_input2.send_keys(otp2)

        otp_input3 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div/div/div/div/div/div/div[2]/div/form/div/div[1]/div/input[3]"))
        )
        otp_input3.send_keys(otp3)

        otp_input4 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div/div/div/div/div/div/div[2]/div/form/div/div[1]/div/input[4]"))
        )
        otp_input4.send_keys(otp4)

        # Find and click the continue button
        continue_button = driver.find_element(By.XPATH, "/html/body/div/div/div/div/div/div/div[2]/div/form/div/div[2]/div/button")
        continue_button.click()

        try:
            SBL_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div/div[2]/div[2]/a/div[1]/img"))
            )
            SBL_button.click()
            print("Clicked on SBL button.")
        except TimeoutException:
            print("SBL button not found. Clicking on Dashboard button instead.")
            Dashboard_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/nav/div[2]/div/div/div/div[2]/ul[1]/a[1]/div[2]/span"))
            )
            Dashboard_button.click()

    except TimeoutException as e:
        print(f"TimeoutException occurred: {e}")
        assert False, "Timeout waiting for element."

    finally:
        # Add any necessary cleanup or verification steps here
        pass

@pytest.mark.parametrize("username, password, otp1, otp2, otp3, otp4, role", [
    ("10001", "123456", "1", "1", "1", "1", 'Admin'),
    # Add more username-password combinations as needed
])
def test_consent_request(setup, username, password, otp1, otp2, otp3, otp4, role):
    driver = setup

    login(driver, username, password, otp1, otp2, otp3, otp4, role)

    driver.get('https://uat-admin.kaabilfinance.com/home')

    global row_counter

    try:
        row_counter += 1
        sheet[f"A{row_counter}"] = "Loan Number"
        sheet[f"B{row_counter}"] = "SL3511291"
        Loan_Search_xpath = "/html/body/div[1]/div/div/header/div/div[1]/span/span/input"
        try:
            Loan_search = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, Loan_Search_xpath))
            )
            Loan_search.send_keys('BD0604549')
            Loan_search.send_keys(Keys.ENTER)
            Loan_search_present = True
        except:
            Loan_search_present = False
            print("Loan Number is Invalid")

        # Check Applicant button
        try:
            Sanction_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//a[contains(@class, "MuiTypography-root") and contains(@class, "MuiLink-root") and contains(@class, "MuiLink-underlineHover") and contains(@class, "MuiTypography-colorPrimary") and text()="Sanction"]'))
            )
            Sanction_button.click()
            Sanction_button_present = True
        except:
            Sanction_button_present = False
            print("Sanction button not found")

        
        try:
            calculate_insurance_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "MuiButtonBase-root") and contains(@class, "MuiButton-root") and contains(@class, "MuiButton-text")]//span[text()="Calculate Insurance"]'))
            )
            calculate_insurance_button_present = True
            row_counter += 1
            sheet[f"A{row_counter}"] = "Calculate Insurance Button"
            sheet[f"B{row_counter}"] = "Available"
        except:
            calculate_insurance_button_present = False
            row_counter += 1
            sheet[f"A{row_counter}"] = "Calculate Insurance Button"
            sheet[f"B{row_counter}"] = "Not Available"            
            print("calculate insurance button not found")

        try:
            recalculate_insurance_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "MuiButtonBase-root") and contains(@class, "MuiButton-root") and contains(@class, "MuiButton-text")]//span[text()="Recalculate Insurance"]'))
            )
            row_counter += 1
            sheet[f"A{row_counter}"] = "Recalculate Insurance Button"
            sheet[f"B{row_counter}"] = "Available"
        except:
            recalculate_insurance_button_present = False
            row_counter += 1
            sheet[f"A{row_counter}"] = "Recalculate Insurance Button"
            sheet[f"B{row_counter}"] = "Not Available"            
            print("recalculate insurance button not found")      




   
    except Exception as e:
        print(f"Exception occurred: {str(e)}")
        # Handle any exceptions that might occur during the WebDriver operations



    input("Press Enter to quit the WebDriver...")

    # Save the Excel file
    wb.save("test_results.xlsx")

    # Print a message
    print("Test results saved to test_results.xlsx")

# Run the tests if this script is executed directly
if __name__ == "__main__":
    row_counter = 25  # Initialize the row_counter
    pytest.main()
