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
        sheet[f"B{row_counter}"] = "BD0604549"
        Loan_Search_xpath = "/html/body/div[1]/div/div/header/div/div[1]/span/span/input"
        try:
            Loan_search = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, Loan_Search_xpath))
            )
            Loan_search.send_keys('BD0604549')
            Loan_search.send_keys(Keys.ENTER)
            Loan_search_present = True
            print("Loan Number is valid")
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
            print("Sanction button found")
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
            print("calculate insurance button found")
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
            print("recalculate insurance button found")
        except:
            recalculate_insurance_button_present = False
            row_counter += 1
            sheet[f"A{row_counter}"] = "Recalculate Insurance Button"
            sheet[f"B{row_counter}"] = "Not Available"            
            print("recalculate insurance button not found")      


        try:
            calculate_insurance_button.click()
            Nominee_name = 'Test Test'
            Nominee_name_search = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'nominee_name'))
            )
            select = Select(Nominee_name_search)
            select.select_by_visible_text(Nominee_name)  # Replace with the desired option text
            row_counter += 1
            sheet[f"A{row_counter}"] = "Nominee Name"
            sheet[f"B{row_counter}"] = Nominee_name
            print("Nominee Name Option found")
        except:
            row_counter += 1
            sheet[f"A{row_counter}"] = "Nominee Name"
            sheet[f"B{row_counter}"] = "Not Available"            
            print("Nominee Name Option not found") 


        try:
            # calculate_insurance_button.click()
            nominee = 'Brother'
            nominee_relationship_search = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'nominee_relationship'))
            )
            select = Select(nominee_relationship_search)
            select.select_by_visible_text(nominee)  # Replace with the desired option text
            row_counter += 1
            sheet[f"A{row_counter}"] = "Nominee Relationship"
            sheet[f"B{row_counter}"] = nominee
            print("Nominee Relationship Option found") 
        except:
            row_counter += 1
            sheet[f"A{row_counter}"] = "nominee_relationship"
            sheet[f"B{row_counter}"] = "Not Available"            
            print("Nominee Relationship Option not found") 



        try:
            # calculate_insurance_button.click()
            income = '200000'
            annual_income_search = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'annual_income'))
            )
            annual_income_search.send_keys(income)
            row_counter += 1
            sheet[f"A{row_counter}"] = "Annual Income"
            sheet[f"B{row_counter}"] = income
            print("Annual Income Option found") 
        except:
            row_counter += 1
            sheet[f"A{row_counter}"] = "Annual Income"
            sheet[f"B{row_counter}"] = "Not Available"            
            print("Annual Income Option not found")



        try:
            # calculate_insurance_button.click()
            save_button = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//span[text()="Save Insurance"]'))
            )
            save_button.click()
            row_counter += 1
            sheet[f"A{row_counter}"] = "Save Insurance Button"
            sheet[f"B{row_counter}"] = "Click"
            print("Insurance Button Option found") 
            # Handle the first alert
            try:
                alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert_text = alert.text
                alert.accept()  # or alert.dismiss()
                # assert alert_text == 'Insurance premium saved'
                print(f"First Alert message: {alert_text}")
                
                row_counter += 1
                sheet[f"A{row_counter}"] = "Insurance details save"
                sheet[f"B{row_counter}"] = alert_text
                print("alert found")
            except TimeoutException:
                print("alert not found")
                row_counter += 1
                sheet[f"A{row_counter}"] = "Insurance details save"
                sheet[f"B{row_counter}"] = "Failed"
        except:
            row_counter += 1
            sheet[f"A{row_counter}"] = "Save Insurance Button"
            sheet[f"B{row_counter}"] = "Not Available"            
            print("Save Insurance Button Option not found")

        try:
# again redirect to the scention page
            Sanction_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//a[contains(@class, "MuiTypography-root") and contains(@class, "MuiLink-root") and contains(@class, "MuiLink-underlineHover") and contains(@class, "MuiTypography-colorPrimary") and text()="Sanction"]'))
            )
            Sanction_button.click()
            Sanction_button_present = True
            print("Sanction button found")

            edit_link = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//th[@class="MuiTableCell-root MuiTableCell-head jss47 MuiTableCell-alignRight"]//a[text()="Edit"]'))
            )

            # Click the <a> element
            edit_link.click()

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.letter_loan_id'))
            )
            loan_id_value = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "Letter Loan Number"
            sheet[f"B{row_counter}"] = loan_id_value

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.amount'))
            )
            loan_amount = input_element.get_attribute('value')  

            row_counter += 1
            sheet[f"A{row_counter}"] = "Loan Request Amount"
            sheet[f"B{row_counter}"] = loan_amount

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.sanction_amount'))
            )
            loan_sanction_amount = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "Sanction Amount"
            sheet[f"B{row_counter}"] = loan_sanction_amount

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.roi'))
            )
            loan_roii = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "ROI"
            sheet[f"B{row_counter}"] = loan_roii

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.pf_rate'))
            )
            loan_pf = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "PF (%)"
            sheet[f"B{row_counter}"] = loan_pf

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.tenure'))
            )
            loan_tenuree = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "Tenure"
            sheet[f"B{row_counter}"] = loan_tenuree

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.insurance_amount'))
            )
            loan_insuranceamount = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "Insurance Amount"
            sheet[f"B{row_counter}"] = loan_insuranceamount

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.insured_applicant'))
            )
            loan_ins_applicant = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "Insured Applicant"
            sheet[f"B{row_counter}"] = loan_ins_applicant

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.stamp_duty'))
            )
            loan_st_duty = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "Stamp Duty"
            sheet[f"B{row_counter}"] = loan_st_duty

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.processing_fees'))
            )
            loan_proc_fees = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "Processing Fees"
            sheet[f"B{row_counter}"] = loan_proc_fees

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.imd_deduction'))
            )
            loan_imd_deduction = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "IMD Deduction"
            sheet[f"B{row_counter}"] = loan_imd_deduction

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.internal_bt_deduction'))
            )
            loan_internal_deduction = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "Internal BT Deduction"
            sheet[f"B{row_counter}"] = loan_internal_deduction

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.other_charges'))
            )
            loan_other_charges = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "Other charges"
            sheet[f"B{row_counter}"] = loan_other_charges

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.other_charges_comment'))
            )
            loan_other_charge_comment = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "Other Charge Comment"
            sheet[f"B{row_counter}"] = loan_other_charge_comment

            # input_element = WebDriverWait(driver, 20).until(
            #     EC.presence_of_element_located((By.NAME, 'loan.letter_loan_id'))
            # )
            # loan_id_value = input_element.get_attribute('value')
            cell = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/main/div[2]/div/div[7]/div/p/div/form/div/table/tbody/tr[15]/td')
            loan_type = cell.text 
            row_counter += 1
            sheet[f"A{row_counter}"] = "Loan Type"
            sheet[f"B{row_counter}"] = loan_type

            input_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'loan.allcloud_file_num'))
            )
            loan_allcloud_file_num = input_element.get_attribute('value')

            row_counter += 1
            sheet[f"A{row_counter}"] = "All cloud File Number"
            sheet[f"B{row_counter}"] = loan_allcloud_file_num

        except Exception as e:
            print(f"exception occurrred: {str(e)}")


     
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
