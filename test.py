import pytest
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

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

        # Wait for the OTP input field to become visible
        otp_input1 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div/div/div/div/div/div/div[2]/div/form/div/div[1]/div/input[1]"))
        )

        otp_input1.send_keys(otp1)

       # Wait for the OTP input field to become visible
        otp_input2 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div/div/div/div/div/div/div[2]/div/form/div/div[1]/div/input[2]"))
        )

        otp_input2.send_keys(otp2)

       # Wait for the OTP input field to become visible
        otp_input3 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div/div/div/div/div/div/div[2]/div/form/div/div[1]/div/input[3]"))
        )

        otp_input3.send_keys(otp3)

       # Wait for the OTP input field to become visible
        otp_input4 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div/div/div/div/div/div/div[2]/div/form/div/div[1]/div/input[4]"))
        )

        otp_input4.send_keys(otp4)

        # # Pause the test execution to allow manual entry of OTP
        # input("Please enter the OTP manually and press Enter to continue...")

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

            # Click on Dashboard button
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
   

   # cfo dashboard diffrent and anlytic all change
     ("10001", "123456", "3", "3", "3", "3", 'Admin'),        # collection report, todays performance,
    #  ("2000", "123456", "1", "1", "1", "1", 'GM'),       # collection report, todays performance, SBL closed loans, SBL DPD Report present ho and collection R me applied Enach not, active enach not present na ho, account r sbl intrest present na ho
    #  ("4444", "123456", "1", "1", "1", "1", 'CFO'),      # collection report, todays performance, collection R me applied enach no present na ho
    #  ("1000001", "123456", "1", "1", "1", "1", 'Sales'),     # only Collection not show and Collection
    #  ("1000002", "123456", "1", "1", "1", "1", 'BCM'),       # only SBL reviw report showing and Collection
    # ("8009", "123456", "1", "1", "1", "1", 'BM'),           # only SBL reviw report showing and Collection
    #  ("1000003", "123456", "1", "1", "1", "1", 'Bpo'),       # only Collection
    #  ("1000004", "123456", "1", "1", "1", "1", 'collection'),  # only Collection
    # # ("1000005", "123456", "1", "1", "1", "1", 'Head office - maker'), #check
    #  ("1000006", "123456", "1", "1", "1", "1", 'Head office - Approver'),
    #  ("1000007", "123456", "1", "1", "1", "1", 'Collendingt Maker'),
    #  ("1000008", "123456", "1", "1", "1", "1", 'Colleanding Approver'),
    # ("1000009", "123456", "1", "1", "1", "1", 'Ho Credit Maker'),
    #  ("1000010", "123456", "1", "1", "1", "1", 'Ho Credit Approver'), #$
    #  ("1000011", "123456", "1", "1", "1", "1", 'Ho Collection Maker'),
    #  ("1000012", "123456", "1", "1", "1", "1", 'Ho Collection Approver'),
#  #  ("1000013", "123456", "1", "1", "1", "1", 'Ho Ops Maker'),
    #  ("1000014", "123456", "1", "1", "1", "1", 'Ho Ops Approver'), # ok
    #  ("1000015", "123456", "1", "1", "1", "1", 'Ho Sales Approver'),#ok
    #  ("1000016", "123456", "1", "1", "1", "1", 'Cluster Sale Manager'), #ok
    #  ("1000017", "123456", "1", "1", "1", "1", 'Cluster Credit Manager'), #ok
    #  ("1000018", "123456", "1", "1", "1", "1", 'RSM/Region Sales Manager'),
    #  ("1000019", "123456", "1", "1", "1", "1", 'ZSM/Zonal Sale Manager'),
    #  ("1000020", "123456", "1", "1", "1", "1", 'Executive Assistane Of ZSM'),
    #  ("1000021", "123456", "1", "1", "1", "1", 'KTR/Kaabil KTR'),
    #  ("1000022", "123456", "1", "1", "1", "1", 'ASM/Area Sale Manager'),
    #  ("1000023", "123456", "1", "1", "1", "1", 'ACM/Area Credit Manager'),
    #  ("1000024", "123456", "1", "1", "1", "1", 'Zonal Credit Manager'),
    #  ("1000025", "123456", "1", "1", "1", "1", 'Customer Care'),
    # ("1000026", "123456", "1", "1", "1", "1", 'Account Maker'), #check
    # ("1000027", "123456", "1", "1", "1", "1", 'Account Checker'),#check
    #  ("1000028", "123456", "1", "1", "1", "1", 'Legal Manager'),
    #  ("1000029", "123456", "1", "1", "1", "1", 'Collection Manager'),
    #  ("1000033", "123456", "1", "1", "1", "1", 'HR'),
    #  ("1000034", "123456", "1", "1", "1", "1", 'IT Manager'),
    #  ("1000035", "123456", "1", "1", "1", "1", 'Channel Partner'),
    #  ("1000036", "123456", "1", "1", "1", "1", 'Sale Lead Generator'),
    #  ("1000037", "123456", "1", "1", "1", "1", 'VP- Corporate Strategy'),
    #  ("1000038", "123456", "1", "1", "1", "1", 'PIVG Manager'),

  
    # Add more username-password combinations as needed
])


#  Test Request and dashboard

def test_my_requests(setup, username, password, otp1, otp2, otp3, otp4, role):
    driver = setup

    login(driver, username, password, otp1, otp2, otp3, otp4, role)


    driver.get('https://uat-admin.kaabilfinance.com/home')


    # Initialize flags for element presence
    my_requests_text_present = False
    pending_button_present = False
    approved_button_present = False
    rejected_button_present = False
    my_approver_text_present = False
    A_pending_button_present = False

    try:
        # Check if "My Requests" text is present
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//h3[text()='My Requests']"))
            )
            my_requests_text_present = True
        except:
            my_requests_text_present = False
        
        if my_requests_text_present:

            # Check if Pending button is present
            pending_button_xpath = "//div[@style='width: 100px; height: 100px; background-color: rgb(109, 103, 228); border-radius: 10px; color: white; margin: 10px; display: flex; flex-direction: column; align-items: center; justify-content: space-around;']"
            try:
                pending_button = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, pending_button_xpath))
                )
                pending_button_present = True
            except:
                pending_button_present = False

            # Check if Approved button is present
            approved_button_xpath = "//div[@style='width: 100px; height: 100px; background-color: rgb(53, 94, 59); border-radius: 10px; color: white; margin: 10px; display: flex; flex-direction: column; align-items: center; justify-content: space-around;']"
            try:
                approved_button = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, approved_button_xpath))
                )
                approved_button_present = True
            except:
                approved_button_present = False

            # Check if Rejected button is present
            rejected_button_xpath = "//div[@style='width: 100px; height: 100px; background-color: rgb(146, 39, 36); border-radius: 10px; color: white; margin: 10px; display: flex; flex-direction: column; align-items: center; justify-content: space-around;']"
            try:
                rejected_button = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, rejected_button_xpath))
                )
                rejected_button_present = True
            except:
                rejected_button_present = False

        # Check if "My Approvals" text is present
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//h3[text()='My Approvals']"))
            )
            my_approver_text_present = True
        except:
            my_approver_text_present = False

        if my_approver_text_present:

            # Check if A_Pending button is present
            A_pending_button_xpath = "//div[@style='width: 100px; height: 100px; background-color: rgb(0, 51, 102); border-radius: 10px; color: white; margin: 25px auto auto; display: flex; flex-direction: column; align-items: center; justify-content: space-around;']"
            try:
                A_pending_button = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, A_pending_button_xpath))
                )
                A_pending_button_present = True
            except:
                A_pending_button_present = False

    except Exception as e:
        print(f"Exception occurred: {str(e)}")
        # Handle any exceptions that might occur during the WebDriver operations



    # Initialize flags for element presence
    Login_Count_present = False
    Sanction_Letter_Count_present = False
    Colender_Dis_Count_present = False
    Disbursement_Count_present = False
    Colender_Sanctions_Chart_present = False
    Monthly_ENACH_Collection_present = False
    SBL_Cash_Inflow_present = False
    Collection_NACH_Status_present = False
    SBL_Loan_Under_Process_present = False

    try:
        # Login Count Present
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//b[text()='Login Count ']"))
            )
            Login_Count_present = True
        except:
            Login_Count_present = False

        # Sanction Letter Count
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//b[text()='Sanction Letter Count']"))
            )
            Sanction_Letter_Count_present = True
        except:
            Sanction_Letter_Count_present = False

        # Colender Disbursement Count
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//b[text()='Colender Disbursement Count']"))
            )
            Colender_Dis_Count_present = True
        except:
            Colender_Dis_Count_present = False

        # Disbursement Count
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//b[text()='Disbursement Count']"))
            )
            Disbursement_Count_present = True
        except:
            Disbursement_Count_present = False

        # Colender's Sanctions Chart
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//b[text()='Colender's Sanctions ']"))
            )
            Colender_Sanctions_Chart_present = True
        except:
            Colender_Sanctions_Chart_present = False

        # Monthly ENACH Collection Success
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//b[text()='Monthly ENACH Collection Success (%)']"))
            )
            Monthly_ENACH_Collection_present = True
        except:
            Monthly_ENACH_Collection_present = False

        # SBL Cash Inflow
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//b[text()='SBL Cash Inflow ']"))
            )
            SBL_Cash_Inflow_present = True
        except:
            SBL_Cash_Inflow_present = False

        # Collection NACH Status
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//b[text()='Collection NACH Status']"))
            )
            Collection_NACH_Status_present = True
        except:
            Collection_NACH_Status_present = False

        # SBL Loan Under Process
        try:
            element = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//b[text()='SBL Loan Under Process']"))
            )
            SBL_Loan_Under_Process_present = True
        except:
            SBL_Loan_Under_Process_present = False

    except Exception as e:
        print(f"Exception occurred: {str(e)}")
        # Handle any exceptions that might occur during the WebDriver operations



       # Navigate to the customers page
    # driver.get('https://uat-admin.kaabilfinance.com/home')  # Replace with your actual URL

    # for customers
    # 
    Customer_Button_present = False
    Add_Customer_Button_present = False

    try:
        # Check if Customer Option is available
        Customer_Button = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[contains(text(),'Customer')]"))
        )
        Customer_Button_present = True
        Customer_Button.click()

        # Check Add Customer Button
        if Customer_Button_present:
            try:
                Add_Customer_Button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//button/span[contains(text(),'Add Customer')]"))
                )
                Add_Customer_Button_present = True
            except TimeoutException:
                Add_Customer_Button_present = False

    except TimeoutException:
        print("Timed out waiting for Customer button")
        Customer_Button_present = False
        Add_Customer_Button_present = False
    except Exception as e:
        print(f"Exception occurred: {str(e)}")
        Customer_Button_present = False
        Add_Customer_Button_present = False


    # For Approvals Button
    # driver.get('https://uat-admin.kaabilfinance.com/home')


    approvals_button_present = False
    approval_list_button_present = False
    my_requests_button_present = False

    try:
        # Check if Approvals Button is present
        approvals_button = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Approvals']"))
        )
        
        approvals_button_present = True
        approvals_button.click()

        # Check Approvals List Button
        if approvals_button_present:
            try:
                approval_list_button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Approval List']"))
                )
                
                approval_list_button_present = True
            except NoSuchElementException:
                print("Element not found: Approval List button")
            except TimeoutException:
                print("Timed out waiting for Approval List button")

        # Check My Requests Button
            try:
                my_requests_button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='My Requests']"))
                )
                
                my_requests_button_present = True
            except NoSuchElementException:
                print("Element not found: My Requests button")
            except TimeoutException:
                print("Timed out waiting for My Requests button")

    except Exception as e:
        print(f"Exception occurred: {str(e)}")

        # Set all variables to False if an exception occurs
        approvals_button_present = False
        approval_list_button_present = False
        my_requests_button_present = False


  # for test Disbursement


  # Check Disbursement Option is avaliable or Not

    Disbursement_Button_present = False
    Disbursal_Pending_Button_present = False
    Pending_Memo_Button_present = False

    try:
        Disbursement_Button = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Disbursement']"))
        )
        
        Disbursement_Button_present = True
        Disbursement_Button.click()

        # Disbursal Pending Button
        if Disbursement_Button_present:
            try:
                Disbursal_Pending_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Disbursal Pending']"))
                )
                
                Disbursal_Pending_Button_present = True
            except NoSuchElementException:
                print("Element not found: Disbursal Pending button")
            except TimeoutException:
                print("Timed out waiting for Disbursal Pending button")

        # Pending Memo Button   
            try:
                Pending_Memo_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Pending Memo']"))
                )
                
                Pending_Memo_Button_present = True
            except NoSuchElementException:
                print("Element not found: Pending Memo button")
            except TimeoutException:
                print("Timed out waiting for Pending Memo button")

    except Exception as e:
        print(f"Exception occurred: {str(e)}")

        # Set all variables to False if an exception occurs
        Disbursement_Button_present = False
        Disbursal_Pending_Button_present = False
        Pending_Memo_Button_present = False


#     # for SBL Loans Reports

#     # driver.get('https://uat-admin.kaabilfinance.com/home')  

# # Check Disbursement Option is avaliable or Not

    SBL_Loans_Button_present = False
    My_SBL_Task_Button_present = False
    Search_Loans_Button_present = False
    SBL_Files_Button_present = False
    Sales_Stage_Button_present = False
    Credit_Pending_Button_present = False
    Operation_Pending_Button_present = False
    Sanction_Pending_Button_present = False

    try:
        SBL_Loans_Button = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='SBL Loans']"))
        )
        
        SBL_Loans_Button_present = True
        SBL_Loans_Button.click()

        # My SBL Tasks Button
        if SBL_Loans_Button_present:
            try:
                My_SBL_Task_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='My SBL Tasks']"))
                )
                
                My_SBL_Task_Button_present = True
            except NoSuchElementException:
                print("Element not found: My SBL Tasks button")
            except TimeoutException:
                print("Timed out waiting for My SBL Tasks button")

        # Search Loans Button   
            try:
                Search_Loans_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Search Loans']"))
                )
                
                Search_Loans_Button_present = True
            except NoSuchElementException:
                print("Element not found: Search Loans button")
            except TimeoutException:
                print("Timed out waiting for Search Loans button")

        # SBL Files Button
            try:
                SBL_Files_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='SBL Files']"))
                )
                
                SBL_Files_Button_present = True
            except NoSuchElementException:
                print("Element not found: SBL Files button")
            except TimeoutException:
                print("Timed out waiting for SBL Files button")

        # Sales Stage Button   
            try:
                Sales_Stage_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Sales Stage']"))
                )
                
                Sales_Stage_Button_present = True
            except NoSuchElementException:
                print("Element not found: Sales Stage button")
            except TimeoutException:
                print("Timed out waiting for Sales Stage button")

        # Credit Pending Button
            try:
                Credit_Pending_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Credit Pending']"))
                )
                
                Credit_Pending_Button_present = True
            except NoSuchElementException:
                print("Element not found: Credit Pending button")
            except TimeoutException:
                print("Timed out waiting for Credit Pending button")

        # Operation Pending Button   
            try:
                Operation_Pending_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Operation Pending']"))
                )
                
                Operation_Pending_Button_present = True
            except NoSuchElementException:
                print("Element not found: Operation Pending button")
            except TimeoutException:
                print("Timed out waiting for Operation Pending button")

        # Sanction Pending Button   
            try:
                Sanction_Pending_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Sanction Pending']"))
                )
                
                Sanction_Pending_Button_present = True
            except NoSuchElementException:
                print("Element not found: Sanction Pending button")
            except TimeoutException:
                print("Timed out waiting for Sanction Pending button")

    except Exception as e:
        print(f"Exception occurred: {str(e)}")

        # Set all variables to False if an exception occurs
        SBL_Loans_Button_present = False
        My_SBL_Task_Button_present = False
        Search_Loans_Button_present = False
        SBL_Files_Button_present = False
        Sales_Stage_Button_present = False
        Credit_Pending_Button_present = False
        Operation_Pending_Button_present = False
        Sanction_Pending_Button_present = False

# # For Collection Report
 

    Collections_Button_present = False
    Virtual_Accounts_Button_present = False
    Field_Visits_Button_present = False
    Branch_Payments_Button_present = False
    Branch_Collections_Button_present = False
    QR_Payments_Button_present = False
    CASH_Payments_Button_present = False
    Print_Receipts_Button_present = False
    Nach_Payments_Button_present = False
    AU_E_Nach_Button_present = False
    Collect_E_Nach_Button_present = False
    Collect_Mail_Nach_Button_present = False
    ICICI_Enach_List_Button_present = False

    try:
        Collections_Button = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Collections']"))
        )
        Collections_Button_present = True
        Collections_Button.click()
    
        # Virtual Accounts Button
        if Collections_Button_present:
            try:
                Virtual_Accounts_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Virtual Accounts']"))
                )
                
                Virtual_Accounts_Button_present = True
            except NoSuchElementException:
                print("Element not found: Virtual Accounts button")
            except TimeoutException:
                print("Timed out waiting for Virtual Accounts button")

             # Field Visits Button   
            try:
                Field_Visits_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Field Visits']"))
                )
                Field_Visits_Button_present = True
            except NoSuchElementException:
                print("Element not found: Field Visits button")
            except TimeoutException:
                print("Timed out waiting for Field Visits button")

                # Branch Payments Button
            try:
                Branch_Payments_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Branch Payments']"))
                )
                Branch_Payments_Button_present = True
            except NoSuchElementException:
                print("Element not found: Branch Payments button")
            except TimeoutException:
                print("Timed out waiting for Branch Payments button")

                 # Branch Collections Button   
            try:
                Branch_Collections_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Branch Collections']"))
                )
                Branch_Collections_Button_present = True
            except NoSuchElementException:
                print("Element not found: Branch Collections button")
            except TimeoutException:
                print("Timed out waiting for Branch Collections button")

                # QR Payments Button
            try:
                QR_Payments_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='QR Payments']"))
                )
                QR_Payments_Button_present = True
            except NoSuchElementException:
                print("Element not found: QR Payments button")
            except TimeoutException:
                print("Timed out waiting for QR Payments button")

                # CASH Payments Button   
            try:
                CASH_Payments_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='CASH Payments']"))
                )
                CASH_Payments_Button_present = True
            except NoSuchElementException:
                print("Element not found: CASH Payments button")
            except TimeoutException:
                print("Timed out waiting for CASH Payments button")

                # Print Receipts Button
            try:
                Print_Receipts_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Print Receipts']"))
                )
                Print_Receipts_Button_present = True
            except NoSuchElementException:
                print("Element not found: Print Receipts button")
            except TimeoutException:
                print("Timed out waiting for Print Receipts button")

             # Nach Payments Button   
            try:
                Nach_Payments_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div/div/nav/div[2]/div/div/div/div[2]/ul[1]/div[5]/div/div/div/a[8]/div/span"))
                )
                Nach_Payments_Button_present = True
            except NoSuchElementException:
                print("Element not found: Nach Payments button")
            except TimeoutException:
                print("Timed out waiting for Nach Payments button")

             # AU E Nach Button
            try:
                AU_E_Nach_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div/div/nav/div[2]/div/div/div/div[2]/ul[1]/div[5]/div/div/div/a[11]/div/span"))
                )
                AU_E_Nach_Button_present = True
            except NoSuchElementException:
                print("Element not found: AU E Nach button")
            except TimeoutException:
                print("Timed out waiting for AU E Nach button")

                # Collect E Nach Button   
            try:
                Collect_E_Nach_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Collect E-Nach']"))
                )
                Collect_E_Nach_Button_present = True
            except NoSuchElementException:
                print("Element not found: Collect E Nach button")
            except TimeoutException:
                print("Timed out waiting for Collect E Nach button")

             # Collect Mail Nach Button
            try:
                Collect_Mail_Nach_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div/div/nav/div[2]/div/div/div/div[2]/ul[1]/div[5]/div/div/div/a[13]/div/span"))
                )
                Collect_Mail_Nach_Button_present = True
            except NoSuchElementException:
                print("Element not found: Collect Mail Nach button")
            except TimeoutException:
                print("Timed out waiting for Collect Mail Nach button")

                # ICIC Enach Butotn
            try:
                ICICI_Enach_List_Button = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='ICICI ENach']"))
                )
                ICICI_Enach_List_Button_present = True
            except NoSuchElementException:
                print("Element not found: Collect Mail Nach button")
            except TimeoutException:
                print("Timed out waiting for Collect Mail Nach button")
    except Exception as e:
        print(f"Exception occurred: {str(e)}")

        # Set all variables to False if an exception occurs



    # for Cash Management

    Cash_Management_Button_present = False
    Check_In_Button_present = False
    Check_Out_Button_present = False
    Cash_Ledgers_Button_present = False
    Bank_Ledgers_Button_present = False
    Deposit_Memo_Button_present = False

    try:
        Cash_Management_Button = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Cash Management')]"))
        )
        
        Cash_Management_Button_present = True
        Cash_Management_Button.click()

        # Check-In Button
        if Cash_Management_Button_present:
            try:
                Check_In_Button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Check-In')]"))
                )
                
                Check_In_Button_present = True
            except NoSuchElementException:
                print("Element not found: Check-In button")
            except TimeoutException:
                print("Timed out waiting for Check-In button")

        # Check-Out Button  
        if Cash_Management_Button_present:
            try:
                Check_Out_Button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Check-Out')]"))
                )
                
                Check_Out_Button_present = True
            except NoSuchElementException:
                print("Element not found: Check-Out button")
            except TimeoutException:
                print("Timed out waiting for Check-Out button")

        # Cash Ledgers Button
            try:
                Cash_Ledgers_Button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Cash Ledgers')]"))
                )
                
                Cash_Ledgers_Button_present = True
            except NoSuchElementException:
                print("Element not found: Cash Ledgers button")
            except TimeoutException:
                print("Timed out waiting for Cash Ledgers button")

        # Bank Ledgers Button
            try:
                Bank_Ledgers_Button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Bank Ledgers')]"))
                )
                
                Bank_Ledgers_Button_present = True
            except NoSuchElementException:
                print("Element not found: Bank Ledgers button")
            except TimeoutException:
                print("Timed out waiting for Bank Ledgers button")

        # Deposit Memo Button
            try:
                Deposit_Memo_Button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Deposit Memo']"))
                )
                
                Deposit_Memo_Button_present = True
            except NoSuchElementException:
                print("Element not found: Deposit Memo button")
            except TimeoutException:
                print("Timed out waiting for Deposit Memo button")

    except Exception as e:
        print(f"Exception occurred: {str(e)}")


    # For Test Document Verfication


    Document_Verification_button_present = False

    try:
        # Check if Approvals Button is present
        Document_Verification_button = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Document Verification')]"))
        )
        
        Document_Verification_button_present = True
        # approvals_button.click()

    except Exception as e:
        print(f"Exception occurred: {str(e)}")

        # Set all variables to False if an exception occurs
        Document_Verification_button_present = False

        
        # All Cloud Ledger
    All_Cloud_Ledger_button_present = False
    try:
        # Check if Approvals Button is present
        All_Cloud_Ledger_button = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='All Cloud Ledger']"))
        )
        All_Cloud_Ledger_button_present = True
        # approvals_button.click()

    except Exception as e:
        print(f"Exception occurred: {str(e)}")

        # Set all variables to False if an exception occurs
        All_Cloud_Ledger_button_present = False

    # Todays Performance


    Today_Performance_button_present = False

    try:
        # Check if Approvals Button is present
        Today_Performance_button = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()=\"Today's Performance\"]"))
        )
        Today_Performance_button_present = True

    except Exception as e:
        print(f"Exception occurred: {str(e)}")

        # Set all variables to False if an exception occurs
        Today_Performance_button_present = False       


        #Bulk EMI Sync
    
    Bulk_EMI_Sync_button_present = False
    try:
        # Check if Approvals Button is present
        Bulk_EMI_Sync_button = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Bulk EMI Sync']"))
        )
        Bulk_EMI_Sync_button_present = True
        # approvals_button.click()

    except Exception as e:
        print(f"Exception occurred: {str(e)}")

        # Set all variables to False if an exception occurs
        Bulk_EMI_Sync_button_present = False


        # Insurance Bulk Upload
    Insurance_Bulk_Upload_button_present = False
    try:
        # Check if Approvals Button is present
        Insurance_Bulk_Upload_button = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, "//span[@class='MuiTypography-root MuiListItemText-primary MuiTypography-body1 MuiTypography-displayBlock' and text()='Insurance Bulk Upload']"))
        )
        Insurance_Bulk_Upload_button_present = True
        # approvals_button.click()

    except Exception as e:
        print(f"Exception occurred: {str(e)}")

        # Set all variables to False if an exception occurs
        Insurance_Bulk_Upload_button_present = False

        
        #Repoerts
    try:
        Reports_button_present = False
        Request_Reports_button_present = False
        kaabil_ops_button_present = False
        ICICI_Enach_List_Report_present = False



        # Check Reports Button
        Reports_button = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[text()='Reports']"))
        )
        Reports_button_present = True
        Reports_button.click()

        # Check Request Reports Button
        if Reports_button_present:
            try:
                Request_Reports_button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div/div/main/div[2]/div/div/button/span[1]"))
                )
                Request_Reports_button_present = True
                Request_Reports_button.click()

            except TimeoutException:
                print("Request Reports button not found or clickable")
                Request_Reports_button_present = False

        # Check Business Division and Select Kaabil Ops
        if Reports_button_present and Request_Reports_button_present:
            try:
                select_Business_Division = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[1]/div[2]/select'))
                )
                select = Select(select_Business_Division)
                select.select_by_value("kaabil_ops")
                kaabil_ops_button_present = True

            except TimeoutException:
                print("Business Division select element not found or selectable")
                kaabil_ops_button_present = False

        # Check ICICI Enach List Report Type
        if Reports_button_present and kaabil_ops_button_present:
            try:
                ICICI_Enach_List_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(ICICI_Enach_List_Report_Type)
                select.select_by_value("kaabil_icici_enach_list")
                ICICI_Enach_List_Report_present = True

            except TimeoutException:
                print("ICICI Enach List Report Type select element not found or selectable")
                ICICI_Enach_List_Report_present = False



    except Exception as e:
        print(f"Exception occurred: {str(e)}")



    driver.get('https://uat-admin.kaabilfinance.com/home')

    try:
        Reports_button_present = False
        Request_Reports_button_present = False
        Branch_level_Log_Report_Type_present = False
        SBL_Review_Report_Type_present = False
        SBL_RO_Productivity_Report_Type_present = False
        SBL_BT_Loan_Report_Type_present = False
        Loan_Under_Processing_Report_Type_present = False
        Loan_Stage_Report_Type_present = False
        Sanction_Letter_Report_Report_Type_present = False
        SUD_Report_Type_present = False
        Disbursement_Report_Type_present = False
        Disbursement_Payments_Report_Type_present = False
        Disbursement_Partial_Payment_Report_Type_present = False
        Loans_in_Collection_List_Report_Type_present = False
        SBL_Payment_Collected_Report_Type_present = False
        Mobile_Number_NOT_verified_Report_Type_present = False
        Valuation_Report_Type_present = False
        RCU_Report_Type_present = False
        Legal_Report_Type_present = False
        SBL_Closed_Loans_Report_Type_present = False
        SBL_DPD_Report_Type_present = False





        # Check Reports Button
        Reports_button = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[text()='Reports']"))
        )
        Reports_button_present = True
        Reports_button.click()

        # Check Request Reports Button
        if Reports_button_present:
            try:
                Request_Reports_button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div/div/main/div[2]/div/div/button/span[1]"))
                )
                Request_Reports_button_present = True
                Request_Reports_button.click()

            except TimeoutException:
                print("Request Reports button not found or clickable")
                Request_Reports_button_present = False

        # Check Business Division and Select SBL
        if Reports_button_present and Request_Reports_button_present:
            try:
                select_Business_Division = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[1]/div[2]/select'))
                )
                select = Select(select_Business_Division)
                select.select_by_value("sbl")
                SBL_button_present = True

            except TimeoutException:
                print("Business Division select element not found or selectable")
                SBL_button_present = False



     
        
            # Check Sanction Letter Report Type

        if Reports_button_present and SBL_button_present:
            try:
                Branch_level_Log_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Branch_level_Log_Report_Type)
                select.select_by_value("sbl_branch_level_login")
                Branch_level_Log_Report_Type_present = True

            except TimeoutException:
                print("ICICI Enach List Report Type select element not found or selectable")
                Branch_level_Log_Report_Type_present = False

        try:
            # Wait for the Sanction Letter Report Type select element to be visible and interactable
            Sanction_Letter_Report_Report_Type = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
            )
            
            # Initialize Select object for Sanction Letter Report Type
            select_sanction_letter = Select(Sanction_Letter_Report_Report_Type)
            
            # Select by value "sbl_sanction_letters"
            select_sanction_letter.select_by_value("sbl_sanction_letters")
            Sanction_Letter_Report_Report_Type_present = True
            # Optional: Check if the correct option is selected

        except TimeoutException:
            print("Timeout: Sanction Letter Report Type select element not found or selectable")
            Sanction_Letter_Report_Report_Type_present = False
        except NoSuchElementException:
            print("Element not found: Sanction Letter Report Type select element not found on the page")
            Sanction_Letter_Report_Report_Type_present = False
        except Exception as e:
            print(f"Error: {e}")
            Sanction_Letter_Report_Report_Type_present = False
            
        #     # Check Disbursement Report  Report Type

        try:
            # Wait for the SUD Report Type select element to be visible and interactable
            SUD_Report_Type = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
            )
            
            # Initialize Select object for SUD Report Type
            select_sud = Select(SUD_Report_Type)
            
            # Select by value "sbl_sud_report"
            select.select_by_value("sbl_sud_report")
            SUD_Report_Type_present = True

        except TimeoutException:
            print("Timeout: SUD Report Type select element not found or selectable")
            SUD_Report_Type_present = False
        except NoSuchElementException:
            print("Element not found: SUD Report Type select element not found on the page")
            SUD_Report_Type_present = False
        except Exception as e:
            print(f"Error: {e}")
            SUD_Report_Type_present = False
     
     #disbursment report

        # Check Disbursement Report  Report Type
        if Reports_button_present and SBL_button_present:
            try:
                Disbursement_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Disbursement_Report_Type)
                select.select_by_value("sbl_disbursement_report")
                Disbursement_Report_Type_present = True

            except TimeoutException:
                print("Disbursement_Report_Type select element not found or selectable")
                Disbursement_Report_Type_present = False

     
     
     # branch level login

     
        # Check SBL Review Report login Report Type
        if Reports_button_present and SBL_button_present:
            try:
                SBL_Review_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_Review_Report_Type)
                select.select_by_value("sbl_loan_review_report")
                SBL_Review_Report_Type_present = True

            except TimeoutException:
                print("SBL_Review_Report_Type List Report Type select element not found or selectable")
                SBL_Review_Report_Type_present = False

        # Check SBL RO Productivity  Report Type
        if Reports_button_present and SBL_button_present:
            try:
                SBL_RO_Productivity_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_RO_Productivity_Report_Type)
                select.select_by_value("sbl_ro_productivity_report")
                SBL_RO_Productivity_Report_Type_present = True

            except TimeoutException:
                print("SBL_RO_Productivity_Report_Type select element not found or selectable")
                SBL_RO_Productivity_Report_Type_present = False
       
        # Check SBL BT Loan Report  Type
        if Reports_button_present and SBL_button_present:
            try:
                SBL_BT_Loan_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_BT_Loan_Report_Type)
                select.select_by_value("get_bt_loan_report")
                SBL_BT_Loan_Report_Type_present = True

            except TimeoutException:
                print("SBL_BT_Loan_Report_Type select element not found or selectable")
                SBL_BT_Loan_Report_Type_present = False
      
        # Check Loan Under Processing  Report Type
        if Reports_button_present and SBL_button_present:
            try:
                Loan_Under_Processing_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Loan_Under_Processing_Report_Type)
                select.select_by_value("sbl_loan_under_progress")
                Loan_Under_Processing_Report_Type_present = True

            except TimeoutException:
                print("Loan_Under_Processing Report Type select element not found or selectable")
                Loan_Under_Processing_Report_Type_present = False
        
        # Check  Loan Stage Report Report Type
        if Reports_button_present and SBL_button_present:
            try:
                Loan_Stage_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Loan_Stage_Report_Type)
                select.select_by_value("sbl_loan_stage_report")
                Loan_Stage_Report_Type_present = True

            except TimeoutException:
                print("Loan_Stage_Report_Type List Report Type select element not found or selectable")
                Loan_Stage_Report_Type_present = False
        


            except TimeoutException:
                print("Disbursement_Report_Type select element not found or selectable")
                Disbursement_Report_Type = False

        # Check Disbursement Payments  Report Type
        if Reports_button_present and SBL_button_present:
            try:
                Disbursement_Payments_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Disbursement_Payments_Report_Type)
                select.select_by_value("sbl_disbursement_payments")
                Disbursement_Payments_Report_Type_present = True

            except TimeoutException:
                print("Disbursement_Payments_Report_Type select element not found or selectable")
                Disbursement_Payments_Report_Type_present = False

        # Check Disbursement Partial Payment  Report Type
        if Reports_button_present and SBL_button_present:
            try:
                Disbursement_Partial_Payment_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Disbursement_Partial_Payment_Report_Type)
                select.select_by_value("sbl_partial_payment_pending")
                Disbursement_Partial_Payment_Report_Type_present = True

            except TimeoutException:
                print("Disbursement_Partial_Payment_Report_Type select element not found or selectable")
                Disbursement_Partial_Payment_Report_Type_present = False

        # Check Loans in Collection List  Report Type
        if Reports_button_present and SBL_button_present:
            try:
                Loans_in_Collection_List_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Loans_in_Collection_List_Report_Type)
                select.select_by_value("sbl_collection_list")
                Loans_in_Collection_List_Report_Type_present = True

            except TimeoutException:
                print("Loans_in_Collection_List_Report_Type select element not found or selectable")
                Loans_in_Collection_List_Report_Type_present = False

        # Check SBL Payment Collected  Report Type
        if Reports_button_present and SBL_button_present:
            try:
                SBL_Payment_Collected_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_Payment_Collected_Report_Type)
                select.select_by_value("sbl_payment_list")
                SBL_Payment_Collected_Report_Type_present = True

            except TimeoutException:
                print("SBL_Payment_Collected_Report_Type select element not found or selectable")
                SBL_Payment_Collected_Report_Type_present = False
       
        # Check Mobile Number NOT verified  Report Type
        if Reports_button_present and SBL_button_present:
            try:
                Mobile_Number_NOT_verified_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Mobile_Number_NOT_verified_Report_Type)
                select.select_by_value("sbl_unverified_contacts")
                Mobile_Number_NOT_verified_Report_Type_present = True

            except TimeoutException:
                print("Mobile_Number_NOT_verified_Report_Type select element not found or selectable")
                Mobile_Number_NOT_verified_Report_Type_present = False

        # Check Valuation Report Report Type
        if Reports_button_present and SBL_button_present:
            try:
                Valuation_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Valuation_Report_Type)
                select.select_by_value("get_valuation_report")
                Valuation_Report_Type_present = True

            except TimeoutException:
                print("Valuation_Report_Type select element not found or selectable")
                Valuation_Report_Type_present = False

        # Check RCU Report Report Type
        if Reports_button_present and SBL_button_present:
            try:
                RCU_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(RCU_Report_Type)
                select.select_by_value("get_rcu_report")
                RCU_Report_Type_present = True

            except TimeoutException:
                print("RCU_Report_Type select element not found or selectable")
                RCU_Report_Type_present = False

        # Check Legal Report Type
        if Reports_button_present and SBL_button_present:
            try:
                Legal_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Legal_Report_Type)
                select.select_by_value("get_legal_report")
                Legal_Report_Type_present = True

            except TimeoutException:
                print("Legal_Report_Type Type select element not found or selectable")
                Legal_Report_Type_present = False


        # Check SBL Closed Loans Report  Type
        if Reports_button_present and SBL_button_present:
            try:
                SBL_Closed_Loans_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                # Create a Select object once the element is located
                select = Select(SBL_Closed_Loans_Report_Type)
                
                # Select by value using the select object
                select.select_by_value("get_closed_loans")
                
                # Check if the correct option is selected
                assert select.first_selected_option.text == "SBL Closed Loans Report"
                
                # Optionally, set a flag if needed
                SBL_Closed_Loans_Report_Type_present = True
                

            except TimeoutException:
                print("SBL_Closed_Loans_Report_Type select element not found or selectable")
                SBL_Closed_Loans_Report_Type_present = False

        # Check SBL DPD Report Type
        if Reports_button_present and SBL_button_present:
            try:
                SBL_DPD_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                    # Create a Select object once the element is located
                select = Select(SBL_DPD_Report_Type)
                
                # Select by value using the select object
                select.select_by_value("sbl_dpd_report")
                
                # Check if the correct option is selected
                selected_option = select.first_selected_option.text
                assert selected_option == "SBL DPD Report"
                SBL_DPD_Report_Type_present = True

            except TimeoutException:
                print("SBL_DPD_Report_Type select element not found or selectable")
                SBL_DPD_Report_Type_present = False



    except Exception as e:
        print(f"Exception occurred: {str(e)}")
        
    


# Report collection 

    driver.get('https://uat-admin.kaabilfinance.com/home')

    Reports_button_present = False
    Request_Reports_button_present = False
    Field_Visit_Report_Type = False
    Telecalling_PTP_Report_Type = False
    Loans_in_Collection_List_Report_Type = False
    EMI_Collection_details_Report_Type = False
    SBL_EMI_Received_Report_Type = False
    Collection_NOT_assigned_Report_Type = False
    Enach_collection_Report_Type = False
    NACH_Collection_details_NOT_Report_Type = False
    Receipt_Tracking_Report_Type = False
    Field_Collection_Report_Type = False
    Applied_ENach_not_in_Collection_Report_Type_present = False
    Active_ENach_not_in_Collection_Sheet_Report_Type = False



    try:
        # Check Reports Button
        Reports_button = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[text()='Reports']"))
        )
        Reports_button_present = True
        Reports_button.click()

        # Check Request Reports Button
        if Reports_button_present:
            try:
                Request_Reports_button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div/div/main/div[2]/div/div/button/span[1]"))
                )
                Request_Reports_button_present = True
                Request_Reports_button.click()

            except TimeoutException:
                print("Request Reports button not found or clickable")
                Request_Reports_button_present = False



        # Check Business Division and Select Collection
        if Reports_button_present and Request_Reports_button_present:
            try:
                select_Business_Division = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[1]/div[2]/select'))
                )
                select = Select(select_Business_Division)
                select.select_by_value("collection")
                Collection_button_present = True

            except TimeoutException:
                print("Business Division select element not found or selectable")
                Collection_button_present = False

        # Check Field Visit Report  Report Type
        if Reports_button_present and Collection_button_present:
            try:
                Field_Visit_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Field_Visit_Report_Type)
                select.select_by_value("collection_field_visits")
                Field_Visit_Report_Type = True

            except TimeoutException:
                print("Field_Visit Report Type select element not found or selectable")
                Field_Visit_Report_Type = False




        # Check Telecalling PTP Report Report Type
        if Reports_button_present and Collection_button_present:
            try:
                Telecalling_PTP_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Telecalling_PTP_Report_Type)
                select.select_by_value("collection_telecalling_report")
                Telecalling_PTP_Report_Type = True

            except TimeoutException:
                print("Telecalling_PTP_Report_Type List Report Type select element not found or selectable")
                Telecalling_PTP_Report_Type = False

        # Check Loans in Collection List   Report Type
        if Reports_button_present and Collection_button_present:
            try:
                Loans_in_Collection_List_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Loans_in_Collection_List_Report_Type)
                select.select_by_value("sbl_collection_list")
                Loans_in_Collection_List_Report_Type = True

            except TimeoutException:
                print("Loans_in_Collection_List_Report_Type select element not found or selectable")
                Loans_in_Collection_List_Report_Type = False
       
        # Check SBL BT Loan Report  Type
        if Reports_button_present and Collection_button_present:
            try:
                EMI_Collection_details_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(EMI_Collection_details_Report_Type)
                select.select_by_value("sbl_emi_collection_mode")
                EMI_Collection_details_Report_Type = True

            except TimeoutException:
                print("EMI_Collection_details_Report_Type select element not found or selectable")
                EMI_Collection_details_Report_Type = False
      
        # Check Loan Under Processing  Report Type
        if Reports_button_present and Collection_button_present:
            try:
                SBL_EMI_Received_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_EMI_Received_Report_Type)
                select.select_by_value("sbl_emi_received_report")
                SBL_EMI_Received_Report_Type = True

            except TimeoutException:
                print("SBL_EMI_Received_Report_Type select element not found or selectable")
                SBL_EMI_Received_Report_Type = False
        
        # Check  Loan Stage Report Report Type
        if Reports_button_present and Collection_button_present:
            try:
                Collection_NOT_assigned_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Collection_NOT_assigned_Report_Type)
                select.select_by_value("collection_not_assigned")
                Collection_NOT_assigned_Report_Type = True

            except TimeoutException:
                print("Collection_NOT_assigned_Report_Type select element not found or selectable")
                Collection_NOT_assigned_Report_Type = False
        
        # Check Sanction Letter Report Type
        if Reports_button_present and Collection_button_present:
            try:
                Enach_collection_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Enach_collection_Report_Type)
                select.select_by_value("kaabil_enach_collection_report")
                Enach_collection_Report_Type = True

            except TimeoutException:
                print("Enach_collection_Report_Type select element not found or selectable")
                Enach_collection_Report_Type = False

        # Check SUD Report Report Type
        if Reports_button_present and Collection_button_present:
            try:
                NACH_Collection_details_NOT_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(NACH_Collection_details_NOT_Report_Type)
                select.select_by_value("sbl_collection_mode_not_available")
                NACH_Collection_details_NOT_Report_Type = True

            except TimeoutException:
                print("NACH_Collection_details_NOT_Report_Type select element not found or selectable")
                NACH_Collection_details_NOT_Report_Type = False

        # Check Disbursement Report  Report Type
        if Reports_button_present and Collection_button_present:
            try:
                Receipt_Tracking_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Receipt_Tracking_Report_Type)
                select.select_by_value("receipt_tracking_report")
                Receipt_Tracking_Report_Type = True

            except TimeoutException:
                print("Disbursement_Report_Type select element not found or selectable")
                Receipt_Tracking_Report_Type = False

        # Check Disbursement Payments  Report Type
        if Reports_button_present and Collection_button_present:
            try:
                Field_Collection_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Field_Collection_Report_Type)
                select.select_by_value("collection_amount_report")
                Field_Collection_Report_Type = True

            except TimeoutException:
                print("Field_Collection_Report_Type select element not found or selectable")
                Field_Collection_Report_Type = False

        # Check Disbursement Partial Payment  Report Type
        if Reports_button_present and Collection_button_present:
            try:
                Applied_ENach_not_in_Collection_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                # Create a Select object once the element is located
                select = Select(Applied_ENach_not_in_Collection_Report_Type)
                # Select by value using the select object
                select.select_by_value("sbl_enach_success_apply_not_collection")
                assert select.first_selected_option.text == "Applied ENach not in Collection Sheet"
                # Optionally, set a flag if needed
                Applied_ENach_not_in_Collection_Report_Type_present = True

            except TimeoutException:
                print("Applied_ENach_not_in_Collection_Report_Type select element not found or selectable")
                Applied_ENach_not_in_Collection_Report_Type_present = False

        # Check Loans in Collection List  Report Type
        if Reports_button_present and Collection_button_present:
            try:
                Active_ENach_not_in_Collection_Sheet_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Active_ENach_not_in_Collection_Sheet_Report_Type)

                select.select_by_value("sbl_active_enach_not_collection")
                
                assert select.first_selected_option.text == "Active ENach not in Collection Sheet"
                # select = Select(Active_ENach_not_in_Collection_Sheet_Report_Type)
                # select.select_by_value("sbl_active_enach_not_collection")
                Active_ENach_not_in_Collection_Sheet_Report_Type = True

            except TimeoutException:
                print("Active_ENach_not_in_Collection_Sheet_Report_Type select element not found or selectable")
                Active_ENach_not_in_Collection_Sheet_Report_Type = False

       



    except Exception as e:
        print(f"Exception occurred: {str(e)}")


# Report Colending

    driver.get('https://uat-admin.kaabilfinance.com/home')

    Reports_button_present = False
    Request_Reports_button_present = False
    SBL_Colender_Monthly_Payout_Report_Type = False
    CoLending_Report_Type = False



    try:
        # Check Reports Button
        Reports_button = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[text()='Reports']"))
        )
        Reports_button_present = True
        Reports_button.click()

        # Check Request Reports Button
        if Reports_button_present:
            try:
                Request_Reports_button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div/div/main/div[2]/div/div/button/span[1]"))
                )
                Request_Reports_button_present = True
                Request_Reports_button.click()

            except TimeoutException:
                print("Request Reports button not found or clickable")
                Request_Reports_button_present = False



        # Check Business Division and Select Collection
        if Reports_button_present and Request_Reports_button_present:
            try:
                select_Business_Division = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[1]/div[2]/select'))
                )
                select = Select(select_Business_Division)
                select.select_by_value("colending")
                Colending_button_present = True

            except TimeoutException:
                print("Business Division select element not found or selectable")
                Colending_button_present = False

        # Check SBL_Colender_Monthly_Payout  Report Type
        if Reports_button_present and Colending_button_present:
            try:
                SBL_Colender_Monthly_Payout_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_Colender_Monthly_Payout_Report_Type)
                select.select_by_value("colender_sbl_monthly_payout_list")
                SBL_Colender_Monthly_Payout_Report_Type = True

            except TimeoutException:
                print("SBL_Colender_Monthly_Payout_Report_Type select element not found or selectable")
                SBL_Colender_Monthly_Payout_Report_Type = False

        # Check Field Visit Report  Report Type
        if Reports_button_present and Colending_button_present:
            try:
                CoLending_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(CoLending_Report_Type)
                select.select_by_value("kaabil_colending_report")
                CoLending_Report_Type = True

            except TimeoutException:
                print("CoLending_Report_Type select element not found or selectable")
                CoLending_Report_Type = False

    except Exception as e:
        print(f"Exception occurred: {str(e)}")

# Report Legal

    driver.get('https://uat-admin.kaabilfinance.com/home')

    Reports_button_present = False
    Request_Reports_button_present = False
    Court_Case_List_Report_Type = False
    Field_Collection_Report_Type = False



    try:
        # Check Reports Button
        Reports_button = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[text()='Reports']"))
        )
        Reports_button_present = True
        Reports_button.click()

        # Check Request Reports Button
        if Reports_button_present:
            try:
                Request_Reports_button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div/div/main/div[2]/div/div/button/span[1]"))
                )
                Request_Reports_button_present = True
                Request_Reports_button.click()

            except TimeoutException:
                print("Request Reports button not found or clickable")
                Request_Reports_button_present = False



        # Check Business Division and Select Collection
        if Reports_button_present and Request_Reports_button_present:
            try:
                select_Business_Division = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[1]/div[2]/select'))
                )
                select = Select(select_Business_Division)
                select.select_by_value("legal")
                Legal_button_present = True

            except TimeoutException:
                print("Business Division select element not found or selectable")
                Legal_button_present = False

        # Check Court Case List  Report Type
        if Reports_button_present and Legal_button_present:
            try:
                Court_Case_List_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Court_Case_List_Report_Type)
                select.select_by_value("legal_court_cases_list")
                Court_Case_List_Report_Type = True

            except TimeoutException:
                print("Court_Case_List_Report_Type select element not found or selectable")
                Court_Case_List_Report_Type = False

        # Check Field Collection Report  Report Type
        if Reports_button_present and Legal_button_present:
            try:
                Field_Collection_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Field_Collection_Report_Type)
                select.select_by_value("collection_amount_report")
                Field_Collection_Report_Type = True

            except TimeoutException:
                print("Field_Collection_Report_Type select element not found or selectable")
                Field_Collection_Report_Type = False

    except Exception as e:
        print(f"Exception occurred: {str(e)}")


# Test Report Accounting

    driver.get('https://uat-admin.kaabilfinance.com/home')

    Reports_button_present = False
    Request_Reports_button_present = False
    Accounting_button_present = False
    Term_Loan_List_Report_Type = False
    Term_Loan_Interest_Statement_Report_Type = False
    Online_Payments_Report_Type = False
    Cash_Payments_Report_Type = False
    Reconcilation_Report_Report_Type = False
    Enach_collection_Report_Type = False
    Non_reconciled_payments_Report_Type = False
    Virtual_Account_list_Type = False
    SBL_Colender_Monthly_Payout_Report_Type = False
    SBL_Colender_Interest_Statement_Report_Type = False
    Login_Fees_Income_Report_Report_Type = False
    IMD_Fees_Income_Report_Type = False
    Automated_Disbursement_Payment_Report_Type = False
    SBL_Cash_inflow_Report_Type = False
    SBL_Income_Report_Type = False
    Gold_Loan_Income_Report_Type = False
    Gold_Monthwise_Report_Type = False
    Insurance_Report_Type = False
    SBL_Interest_Income_Report_Type_present = False
    A_SBL_Closed_Loans_Report_Type_present = False
    A_SBL_DPD_Report_Type_present = False

    try:
        # Check Reports Button
        Reports_button = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[text()='Reports']"))
        )
        Reports_button_present = True
        Reports_button.click()

        # Check Request Reports Button
        if Reports_button_present:
            try:
                Request_Reports_button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div/div/main/div[2]/div/div/button/span[1]"))
                )
                Request_Reports_button_present = True
                Request_Reports_button.click()

            except TimeoutException:
                print("Request Reports button not found or clickable")
                Request_Reports_button_present = False

        # Check Business Division and Accounting_button_present
        if Reports_button_present and Request_Reports_button_present:
            try:
                select_Business_Division = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[1]/div[2]/select'))
                )
                select = Select(select_Business_Division)
                select.select_by_value("accounting")
                Accounting_button_present = True

            except TimeoutException:
                print("Business Division select element not found or selectable")
                Accounting_button_present = False





        # Check Term Loan List  Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                Term_Loan_List_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Term_Loan_List_Report_Type)
                select.select_by_value("accounting_term_loans_list")
                Term_Loan_List_Report_Type = True

            except TimeoutException:
                print("Term_Loan_List_Report Type select element not found or selectable")
                Term_Loan_List_Report_Type = False




        # Check Term Loan Interest Statement  Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                Term_Loan_Interest_Statement_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Term_Loan_Interest_Statement_Report_Type)
                select.select_by_value("term_loan_expense_report")
                Term_Loan_Interest_Statement_Report_Type = True

            except TimeoutException:
                print("Term Loan Interest Statement  List Report Type select element not found or selectable")
                Term_Loan_Interest_Statement_Report_Type = False

        # Check Online Payments   Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                Online_Payments_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Online_Payments_Report_Type)
                select.select_by_value("kaabil_ecollect_payments")
                Online_Payments_Report_Type = True

            except TimeoutException:
                print("Online Payments  select element not found or selectable")
                Online_Payments_Report_Type = False
       
        # Check Cash Payments  Report  Type
        if Reports_button_present and Accounting_button_present:
            try:
                Cash_Payments_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Cash_Payments_Report_Type)
                select.select_by_value("kaabil_cash_payment")
                Cash_Payments_Report_Type = True

            except TimeoutException:
                print("Cash Payments  select element not found or selectable")
                Cash_Payments_Report_Type = False
      
        # Check Reconcilation Report   Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                Reconcilation_Report_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Reconcilation_Report_Report_Type)
                select.select_by_value("kaabil_reconcilation_list")
                Reconcilation_Report_Report_Type = True

            except TimeoutException:
                print("Reconcilation Report Type select element not found or selectable")
                Reconcilation_Report_Report_Type = False
        
        # Check  Enach collection Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                Enach_collection_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Enach_collection_Report_Type)
                select.select_by_value("kaabil_enach_collection_report")
                Enach_collection_Report_Type = True

            except TimeoutException:
                print("Enach collection report  List Report Type select element not found or selectable")
                Enach_collection_Report_Type = False
        
        # Check Non reconciled payments  Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                Non_reconciled_payments_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Non_reconciled_payments_Report_Type)
                select.select_by_value("kaabil_non_reconcile_report")
                Non_reconciled_payments_Report_Type = True

            except TimeoutException:
                print("Non reconciled payments  select element not found or selectable")
                Non_reconciled_payments_Report_Type = False

        # Check Virtual Account list  Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                Virtual_Account_list_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Virtual_Account_list_Type)
                select.select_by_value("kaabil_virtual_accounts")
                Virtual_Account_list_Type = True

            except TimeoutException:
                print("Virtual Account list  select element not found or selectable")
                Virtual_Account_list_Type = False

        # Check SBL Colender Monthly Payout  Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                SBL_Colender_Monthly_Payout_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_Colender_Monthly_Payout_Report_Type)
                select.select_by_value("colender_sbl_monthly_payout_list")
                SBL_Colender_Monthly_Payout_Report_Type = True

            except TimeoutException:
                print("SBL Colender Monthly Payout  select element not found or selectable")
                SBL_Colender_Monthly_Payout_Report_Type = False

        # Check SBL Colender Interest Statement   Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                SBL_Colender_Interest_Statement_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_Colender_Interest_Statement_Report_Type)
                select.select_by_value("colender_expense_report")
                SBL_Colender_Interest_Statement_Report_Type = True

            except TimeoutException:
                print("SBL Colender Interest Statement  select element not found or selectable")
                SBL_Colender_Interest_Statement_Report_Type = False

        # Check Login Fees Income Report   Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                Login_Fees_Income_Report_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Login_Fees_Income_Report_Report_Type)
                select.select_by_value("login_fees_report")
                Login_Fees_Income_Report_Report_Type = True

            except TimeoutException:
                print("Login Fees Income Report  select element not found or selectable")
                Login_Fees_Income_Report_Report_Type = False

        # Check IMD Fees Income Report   Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                IMD_Fees_Income_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(IMD_Fees_Income_Report_Type)
                select.select_by_value("imd_fees_report")
                IMD_Fees_Income_Report_Type = True

            except TimeoutException:
                print("IMD Fees Income Report  select element not found or selectable")
                IMD_Fees_Income_Report_Type = False

        # Check Automated Disbursement Payment   Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                Automated_Disbursement_Payment_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Automated_Disbursement_Payment_Report_Type)
                select.select_by_value("automated_disbursement_payment_report")
                Automated_Disbursement_Payment_Report_Type = True

            except TimeoutException:
                print("Automated Disbursement Payment  select element not found or selectable")
                Automated_Disbursement_Payment_Report_Type = False
       
        # Check SBL Cash inflow report   Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                SBL_Cash_inflow_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_Cash_inflow_Report_Type)
                select.select_by_value("sbl_cash_inflow_report")
                SBL_Cash_inflow_Report_Type = True

            except TimeoutException:
                print("SBL Cash inflow report  select element not found or selectable")
                SBL_Cash_inflow_Report_Type = False

        # Check SBL Income  Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                SBL_Income_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_Income_Report_Type)
                select.select_by_value("sbl_income_statement")
                SBL_Income_Report_Type = True

            except TimeoutException:
                print("SBL Income Report  select element not found or selectable")
                SBL_Income_Report_Type = False

        # Check Gold Loan Income  Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                Gold_Loan_Income_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Gold_Loan_Income_Report_Type)
                select.select_by_value("gold_income_expense_report")
                Gold_Loan_Income_Report_Type = True

            except TimeoutException:
                print("Gold Loan Income Report  select element not found or selectable")
                Gold_Loan_Income_Report_Type = False

        # Check Gold Monthwise Report  Type
        if Reports_button_present and Accounting_button_present:
            try:
                Gold_Monthwise_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Gold_Monthwise_Report_Type)
                select.select_by_value("gold_account_report_month_wise")
                Gold_Monthwise_Report_Type = True

            except TimeoutException:
                print("Gold Monthwise Report  Type select element not found or selectable")
                Gold_Monthwise_Report_Type = False


        # Check Insurance Report   Type
        if Reports_button_present and Accounting_button_present:
            try:
                Insurance_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(Insurance_Report_Type)
                select.select_by_value("create_insurance_report")
                Insurance_Report_Type = True

            except TimeoutException:
                print("Insurance Report  select element not found or selectable")
                Insurance_Report_Type = False

        # Check SBL Interest Income Report  Type
        if Reports_button_present and Accounting_button_present:
            try:
                SBL_Interest_Income_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_Interest_Income_Report_Type)
                select.select_by_value("sbl_income_report")
                SBL_Interest_Income_Report_Type_present = True

            except TimeoutException:
                print("SBL Interest Income Report element not found or selectable")
                SBL_Interest_Income_Report_Type_present = False


        # Check SBL Closed Loans Report  Type
        if Reports_button_present and Accounting_button_present:
            try:
                SBL_Closed_Loans_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_Closed_Loans_Report_Type)
                select.select_by_value("get_closed_loans")
                A_SBL_Closed_Loans_Report_Type_present = True

            except TimeoutException:
                print("SBL Closed Loans Report  select element not found or selectable")
                A_SBL_Closed_Loans_Report_Type_present = False

        # Check SBL DPD Report Type
        if Reports_button_present and Accounting_button_present:
            try:
                SBL_DPD_Report_Type = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div[3]/form/div[2]/div[2]/select'))
                )
                select = Select(SBL_DPD_Report_Type)
                select.select_by_value("sbl_dpd_report")
                A_SBL_DPD_Report_Type_present = True

            except TimeoutException:
                print("SBL DPD Report element not found or selectable")
                A_SBL_DPD_Report_Type_present = False



    except Exception as e:
        print(f"Exception occurred: {str(e)}")



    # Write results to Excel file
    global row_counter
    
    # Write results to Excel file
    # global row_counter
    row_counter += 1

    # Select Role 
    sheet[f"A{row_counter}"] = "Role Name"
    sheet[f"B{row_counter}"] = role

    row_counter += 1

    sheet[f"A{row_counter}"] = "DashBoard"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1

    sheet[f"A{row_counter}"] = "My Requests Text Presence"
    sheet[f"B{row_counter}"] = "Present" if my_requests_text_present else "Not Present"

    row_counter += 1

    sheet[f"A{row_counter}"] = "Pending Button Presence"
    sheet["B5"] = "Present" if pending_button_present else "Not Present"

    row_counter += 1

    sheet[f"A{row_counter}"] = "Approved Button Presence"
    sheet[f"B{row_counter}"] = "Present" if approved_button_present else "Not Present"

    row_counter += 1

    sheet[f"A{row_counter}"] = "Rejected Button Presence"
    sheet[f"B{row_counter}"] = "Present" if rejected_button_present else "Not Present"

    row_counter += 1

    sheet[f"A{row_counter}"] = "My Approver Section Presence"
    sheet[f"B{row_counter}"] = "Present" if my_approver_text_present else "Not Present"

    row_counter += 1

    sheet[f"A{row_counter}"] = "Pending Button Presence"
    sheet[f"B{row_counter}"] = "Present" if A_pending_button_present else "Not Present"

    row_counter += 1

    sheet[f"A{row_counter}"] = "Analytics"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1


    sheet[f"A{row_counter}"] = "Login Count Chart Presence"
    sheet[f"B{row_counter}"] = "Present" if Login_Count_present else "Not Present"

    row_counter += 1

    sheet[f"A{row_counter}"] = "Sanction Letter Count Chart Presence"
    sheet[f"B{row_counter}"] = "Present" if Sanction_Letter_Count_present else "Not Present"   
  
    row_counter += 1

    sheet[f"A{row_counter}"] = "Colender Disbursement Count Chart Presence"
    sheet[f"B{row_counter}"] = "Present" if Colender_Dis_Count_present else "Not Present" 

    row_counter += 1

    sheet[f"A{row_counter}"] = "Disbursement Count Chart Presence"
    sheet[f"B{row_counter}"] = "Present" if Disbursement_Count_present else "Not Present"

    row_counter += 1

    sheet[f"A{row_counter}"] = "Colender's Sanctions Chart Presence"
    sheet[f"B{row_counter}"] = "Present" if Colender_Sanctions_Chart_present else "Not Present" 
    
    row_counter += 1

    sheet[f"A{row_counter}"] = "Monthly ENACH Collection Success(%) Chart Presence"
    sheet[f"B{row_counter}"] = "Present" if Monthly_ENACH_Collection_present else "Not Present"   
  
    row_counter += 1

    sheet[f"A{row_counter}"] = "SBL Cash Inflow Chart Presence"
    sheet[f"B{row_counter}"] = "Present" if SBL_Cash_Inflow_present else "Not Present" 
    
    row_counter += 1

    sheet[f"A{row_counter}"] = "Collection NACH Status Chart Presence"
    sheet[f"B{row_counter}"] = "Present" if Collection_NACH_Status_present else "Not Present"  
    
    row_counter += 1

    sheet[f"A{row_counter}"] = "SBL Loan Under Process Chart Presence"
    sheet[f"B{row_counter}"] = "Present" if SBL_Loan_Under_Process_present else "Not Present"  

    # for Customers
    row_counter += 1

    sheet[f"A{row_counter}"] = "Customers"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1

    sheet[f"A{row_counter}"] = "Customers Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Customer_Button_present else "Not Present"

    row_counter += 1

    sheet[f"A{row_counter}"] = "Add Customer Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Add_Customer_Button_present else "Not Present"
    


    # for approvel button test
    row_counter += 1
    sheet[f"A{row_counter}"] = "Approvals"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Approvals Button Presence"
    sheet[f"B{row_counter}"] = "Present" if approvals_button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Approval List Button Presence"
    sheet[f"B{row_counter}"] = "Present" if approval_list_button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "My Requests Button Presence"
    sheet[f"B{row_counter}"] = "Present" if my_requests_button_present else "Not Present"
   



    # Write results for Disbursement test

    row_counter += 1
    sheet[f"A{row_counter}"] = "Disbursement"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Disbursement Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Disbursement_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Disbursal Pending Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Disbursal_Pending_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Pending Memo  Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Pending_Memo_Button_present else "Not Present"
    

    # for SBL Loans stage

    row_counter += 1    
    sheet[f"A{row_counter}"] = "SBL Loans"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL Loans Button Presence"
    sheet[f"B{row_counter}"] = "Present" if SBL_Loans_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "My SBL Task Button Presence"
    sheet[f"B{row_counter}"] = "Present" if My_SBL_Task_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Search Loans  Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Search_Loans_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL Files Button Presence"
    sheet[f"B{row_counter}"] = "Present" if SBL_Files_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Sales Stage Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Sales_Stage_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Credit Pending Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Credit_Pending_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Operation Pending Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Operation_Pending_Button_present else "Not Present"

    row_counter += 1       
    sheet[f"A{row_counter}"] = "Sanction Pending Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Sanction_Pending_Button_present else "Not Present"


     # Write results for Collection 
 
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collections"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Collections Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Collections_Button_present else "Not Present"
    
    row_counter += 1
    sheet[f"A{row_counter}"] = "Virtual Accounts Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Virtual_Accounts_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Field Visits Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Field_Visits_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Branch Payments Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Branch_Payments_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Branch Collections Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Branch_Collections_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "QR Payments Button Presence"
    sheet[f"B{row_counter}"] = "Present" if QR_Payments_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "CASH Payments Button Presence"
    sheet[f"B{row_counter}"] = "Present" if CASH_Payments_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Print Receipts Buttont Presence"
    sheet[f"B{row_counter}"] = "Present" if Print_Receipts_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Nach Payments Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Nach_Payments_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "AU E Nach Button Presence"
    sheet[f"B{row_counter}"] = "Present" if AU_E_Nach_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Collect E Nach Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Collect_E_Nach_Button_present else "Not Present"

    row_counter += 1       
    sheet[f"A{row_counter}"] = "Collect Mail Nach Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Collect_Mail_Nach_Button_present else "Not Present"    
      
   
   # for Cash Management

    row_counter += 1
    sheet[f"A{row_counter}"] = "Cash Management"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Cash Management Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Cash_Management_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Check In Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Check_In_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Check Out Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Check_Out_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Cash Ledgers Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Cash_Ledgers_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Bank Ledgers Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Bank_Ledgers_Button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Deposit Memo Button Presence"
    sheet[f"B{row_counter}"] = "Present" if Deposit_Memo_Button_present else "Not Present"
 

# for Document Verification

    row_counter += 1
    sheet[f"A{row_counter}"] = "Document Verification"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Document Verification button Presence"
    sheet[f"B{row_counter}"] = "Present" if Document_Verification_button_present else "Not Present"


# for All Cloud Ledger

    row_counter += 1
    sheet[f"A{row_counter}"] = "All Cloud Ledger"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "All Cloud Ledger button Presence"
    sheet[f"B{row_counter}"] = "Present" if All_Cloud_Ledger_button_present else "Not Present"


    # Todays Performance resutl save

    row_counter += 1
    sheet[f"A{row_counter}"] = "Today's Performance"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Today Performance button Presence"
    sheet[f"B{row_counter}"] = "Present" if Today_Performance_button_present else "Not Present"



# Bulk Emi Sync Data save
    row_counter += 1
    sheet[f"A{row_counter}"] = "Bulk EMI Sync"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Bulk EMI Sync button Presence"
    sheet[f"B{row_counter}"] = "Present" if Bulk_EMI_Sync_button_present else "Not Present"


# Insurance Bulk Upload data save
    row_counter += 1
    sheet[f"A{row_counter}"] = "Insurance Bulk Upload"
    sheet[f"B{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Insurance Bulk Upload button Presence"
    sheet[f"B{row_counter}"] = "Present" if Insurance_Bulk_Upload_button_present else "Not Present"

    row_counter += 1


# Report data save SBL and Kabil Ops
    row_counter += 1
    sheet[f"A{row_counter}"] = "Reports"
    sheet[f"B{row_counter}"] = "Result"
   
    row_counter += 1
    sheet[f"A{row_counter}"] = "Report button Presence"
    sheet[f"B{row_counter}"] = "Present" if Reports_button_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Request Reports button Presence"
    sheet[f"B{row_counter}"] = "Present" if Request_Reports_button_present else "Not Present"
    sheet[f"C{row_counter}"] = ""

    row_counter += 1
    sheet[f"A{row_counter}"] = "Business Division"
    sheet[f"B{row_counter}"] = "Report Type"
    sheet[f"C{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Kaabil Ops Type"
    sheet[f"B{row_counter}"] = "ICICI Enach List"
    sheet[f"C{row_counter}"] = "Present" if ICICI_Enach_List_Report_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "Business Division"
    sheet[f"B{row_counter}"] = "Report Type"
    sheet[f"C{row_counter}"] = "Result"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Branch Level Logins "
    sheet[f"C{row_counter}"] = "Present" if Branch_level_Log_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "SBL Review Report "
    sheet[f"C{row_counter}"] = "Present" if SBL_Review_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = " SBL RO Productivity Report"
    sheet[f"C{row_counter}"] = "Present" if SBL_RO_Productivity_Report_Type_present else "Not Present"

    row_counter += 1      
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "SBL BT Loan Report"
    sheet[f"C{row_counter}"] = "Present" if SBL_BT_Loan_Report_Type_present else "Not Present"

    row_counter += 1    
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Loan Under Processing"
    sheet[f"C{row_counter}"] = "Present" if Loan_Under_Processing_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Loan Stage Report"
    sheet[f"C{row_counter}"] = "Present" if Loan_Stage_Report_Type_present else "Not Present"

    row_counter += 1            
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "SUD Report"
    sheet[f"C{row_counter}"] = "Present" if SUD_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Disbursement Report"
    sheet[f"C{row_counter}"] = "Present" if Disbursement_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Disbursement Payments"
    sheet[f"C{row_counter}"] = "Present" if Disbursement_Payments_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Disbursement Partial Payment"
    sheet[f"C{row_counter}"] = "Present" if Disbursement_Partial_Payment_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Loans in Collection List"
    sheet[f"C{row_counter}"] = "Present" if Loans_in_Collection_List_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "SBL Payment Collected"
    sheet[f"C{row_counter}"] = "Present" if SBL_Payment_Collected_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Mobile Number NOT verified"
    sheet[f"C{row_counter}"] = "Present" if Mobile_Number_NOT_verified_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Valuation Report"
    sheet[f"C{row_counter}"] = "Present" if Valuation_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "RCU Report"
    sheet[f"C{row_counter}"] = "Present" if RCU_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Legal Report"
    sheet[f"C{row_counter}"] = "Present" if Legal_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "SBL Closed Loans Report "
    sheet[f"C{row_counter}"] = "Present" if SBL_Closed_Loans_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "SBL DPD Report"
    sheet[f"C{row_counter}"] = "Present" if SBL_DPD_Report_Type_present else "Not Present"

    row_counter += 1
    sheet[f"A{row_counter}"] = "SBL"
    sheet[f"B{row_counter}"] = "Sanction Letter Report"
    sheet[f"C{row_counter}"] = "Present" if Sanction_Letter_Report_Report_Type_present else "Not Present"
        


# Reports data save Collection
    row_counter += 1
    sheet[f"A{row_counter}"] = "Business Division"
    sheet[f"B{row_counter}"] = "Report Type"
    sheet[f"C{row_counter}"] = "Result"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "Field Visit Report"
    sheet[f"C{row_counter}"] = "Present" if Field_Visit_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "Telecalling PTP Report"
    sheet[f"C{row_counter}"] = "Present" if Telecalling_PTP_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "Loans in Collection List"
    sheet[f"C{row_counter}"] = "Present" if Loans_in_Collection_List_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "EMI Collection details"
    sheet[f"C{row_counter}"] = "Present" if EMI_Collection_details_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "SBL EMI Received report "
    sheet[f"C{row_counter}"] = "Present" if SBL_EMI_Received_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "Collection NOT assigned report"
    sheet[f"C{row_counter}"] = "Present" if Collection_NOT_assigned_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "Enach collection report"
    sheet[f"C{row_counter}"] = "Present" if Enach_collection_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "NACH Collection details NOT"
    sheet[f"C{row_counter}"] = "Present" if NACH_Collection_details_NOT_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "Receipt Tracking Report"
    sheet[f"C{row_counter}"] = "Present" if Receipt_Tracking_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "Field Collection Report"
    sheet[f"C{row_counter}"] = "Present" if Field_Collection_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "Applied ENach not in Collection"
    sheet[f"C{row_counter}"] = "Present" if Applied_ENach_not_in_Collection_Report_Type_present else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Collection"
    sheet[f"B{row_counter}"] = "Active ENach not in Collection Sheet"
    sheet[f"C{row_counter}"] = "Present" if Active_ENach_not_in_Collection_Sheet_Report_Type else "Not Present"



# Report Colending data save
    row_counter += 1
    sheet[f"A{row_counter}"] = "Business Division"
    sheet[f"B{row_counter}"] = "Report Type"
    sheet[f"C{row_counter}"] = "Result"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Colending"
    sheet[f"B{row_counter}"] = "SBL Colender Monthly Payout"
    sheet[f"C{row_counter}"] = "Present" if SBL_Colender_Monthly_Payout_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Colending"
    sheet[f"B{row_counter}"] = "CoLending Report"
    sheet[f"C{row_counter}"] = "Present" if CoLending_Report_Type else "Not Present"



# Report Legal
    row_counter += 1
    sheet[f"A{row_counter}"] = "Business Division"
    sheet[f"B{row_counter}"] = "Report Type"
    sheet[f"C{row_counter}"] = "Result"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Legal"
    sheet[f"B{row_counter}"] = "Court Case List"
    sheet[f"C{row_counter}"] = "Present" if Court_Case_List_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Legal"
    sheet[f"B{row_counter}"] = "Field Collection Report"
    sheet[f"C{row_counter}"] = "Present" if Field_Collection_Report_Type else "Not Present"


# Report Accounting
    row_counter += 1
    sheet[f"A{row_counter}"] = "Business Division"
    sheet[f"B{row_counter}"] = "Report Type"
    sheet[f"C{row_counter}"] = "Result"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Term Loan List"
    sheet[f"C{row_counter}"] = "Present" if Term_Loan_List_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Term Loan Interest Statement"
    sheet[f"C{row_counter}"] = "Present" if Term_Loan_Interest_Statement_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Online Payments"
    sheet[f"C{row_counter}"] = "Present" if Online_Payments_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Cash Payments"
    sheet[f"C{row_counter}"] = "Present" if Cash_Payments_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Reconcilation Report"
    sheet[f"C{row_counter}"] = "Present" if Reconcilation_Report_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Enach collection report"
    sheet[f"C{row_counter}"] = "Present" if Enach_collection_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Non reconciled payments"
    sheet[f"C{row_counter}"] = "Present" if Non_reconciled_payments_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Virtual Account list"
    sheet[f"C{row_counter}"] = "Present" if Virtual_Account_list_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "SBL Colender Monthly Payout"
    sheet[f"C{row_counter}"] = "Present" if SBL_Colender_Monthly_Payout_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "SBL Colender Interest Statement"
    sheet[f"C{row_counter}"] = "Present" if SBL_Colender_Interest_Statement_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Login Fees Income Report"
    sheet[f"C{row_counter}"] = "Present" if Login_Fees_Income_Report_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "IMD Fees Income Report"
    sheet[f"C{row_counter}"] = "Present" if IMD_Fees_Income_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Automated Disbursement Payment"
    sheet[f"C{row_counter}"] = "Present" if Automated_Disbursement_Payment_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "SBL Cash inflow report"
    sheet[f"C{row_counter}"] = "Present" if SBL_Cash_inflow_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "SBL Income Report"
    sheet[f"C{row_counter}"] = "Present" if SBL_Income_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Gold Loan Income Report"
    sheet[f"C{row_counter}"] = "Present" if Gold_Loan_Income_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Gold Monthwise Report"
    sheet[f"C{row_counter}"] = "Present" if Gold_Monthwise_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "Insurance Report"
    sheet[f"C{row_counter}"] = "Present" if Insurance_Report_Type else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "SBL Interest Income Report"
    sheet[f"C{row_counter}"] = "Present" if SBL_Interest_Income_Report_Type_present else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "SBL Closed Loans Report"
    sheet[f"C{row_counter}"] = "Present" if A_SBL_Closed_Loans_Report_Type_present else "Not Present"
    row_counter += 1
    sheet[f"A{row_counter}"] = "Accounting"
    sheet[f"B{row_counter}"] = "SBL DPD Report"
    sheet[f"C{row_counter}"] = "Present" if A_SBL_DPD_Report_Type_present else "Not Present"






    

    # Save the Excel file
    wb.save("test_results.xlsx")

    # Print a message
    print("Test results saved to test_results.xlsx")

 

    driver.get('https://uat-admin.kaabilfinance.com/home')

    logout_button = False
    try:
        logout_button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/nav/div[2]/div/div/div/div[2]/ul[3]/li/button/span[1]"))
        )
        logout_button.click()
        logout_button = True
        print("Clicked on the Logout button.")
    except Exception as e:
        print(f"Exception occurred during logout: {e}")

    print("logout 1 success")

    if logout_button == False:   
        try:
            logout_button = WebDriverWait(driver, 30).until(
                logout_button = driver.find_element(By.LINK_TEXT, "Logout")
            )
            logout_button.click()
            print("Clicked on the Logout button.")
        except Exception as e:
            print(f"Exception occurred during logout: {e}")

        print("logout success")
 

    # try:
    #     logout_button = WebDriverWait(driver, 10).until(
    #         EC.presence_of_element_located((By.XPATH, "//span[text()='Logout']/ancestor::button"))
    #     )
    #     logout_button.click()
    #     print("Clicked on the Logout button.")
    # except Exception as e:
    #     print(f"Exception occurred during logout: {e}")





# Run the tests if this script is executed directly
if __name__ == "__main__":
    row_counter = 25  # Initialize the row_counter
    pytest.main()

