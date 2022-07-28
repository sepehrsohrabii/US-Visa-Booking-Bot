from time import sleep
from datetime import datetime, timedelta
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

#s = Service('../../Desktop/chromedriver_win32/chromedriver.exe')
driver = webdriver.Chrome()

# Sign out from account  
def sign_out():
    driver.get("https://ais.usvisa-info.com/en-ae/niv/users/sign_out")
    driver.delete_all_cookies()

# Refresh account to stay login
def refresh():
    # Change City
    try:
        dubai()
        sleep(5)
        abu()
    except:
        print("refresh failed")
        sign_out()
        login()

# Change city to Abu Dhabi
def abu():
    city = driver.find_element(By.ID, "appointments_consulate_appointment_facility_id")
    city1 = Select(city)
    city1.select_by_value("49")        

# Change city to Dubai
def dubai():
    city = driver.find_element(By.ID, "appointments_consulate_appointment_facility_id")
    city1 = Select(city)
    city1.select_by_value("50")

 # login to my account                   
def login():
    global user
    global password
    global link
    global CAPTION
    global TEXT
    
    for h in range(3):
        
        # Set your username here
        user = "########"
        
        # Set your password here
        password = "############"
        
        # Set your account Link here
        link = "https://ais.usvisa-info.com/en-ae/niv/schedule/XXXXXXX/appointment"
        
        for i in range(5):
            driver.get(link)
            try:
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "button[class='ui-button ui-corner-all ui-widget']"))).click()
                
            except selenium.common.exceptions.NoSuchElementException:
                TEXT = "Login PopUP Failed for app: " + user
                print(TEXT)
                
                driver.delete_all_cookies()
                continue
            break
        
        driver.find_element(By.ID, "user_email").send_keys(user)
        driver.find_element(By.ID, "user_password").send_keys(password)
        driver.find_element(By.CSS_SELECTOR,
                            "#new_user > div.radio-checkbox-group.margin-top-30 > label > div").click()

        driver.find_element(By.CSS_SELECTOR, "input[value='Sign In']").click()
        
        # Change City
        try:
            sleep(1)
            abu()
        except:
            TEXT = "Change City failed, app " + user
            print(TEXT)

        break
    
login()      

def visa():
    global now
    global Date
    
    # Read First Date found by Fake Account
    date = open('FAD.txt', 'r')
    Date = date.read()
    day = int(Date[0:2])
    
    # Check month of open date
    if "April 2022" in Date:
        mon = "April 2022"
        min = 20
        max = 30
        if min <= day <= max:
            pass
        else:
            return
    elif "May 2022" in Date:
        mon = "May 2022"
        min = 1
        max = 31
        if min <= day <= max:
            pass
        else:
            return
    elif "June 2022" in Date:
        mon = "June 2022"
        min = 1
        max = 15
        if min <= day <= max:
            pass
        else:
            return
    elif "July 2022" in Date:
        mon = "July 2022"
        min = 1
        max = 15
        if min <= day <= max:
            pass
        else:
            return
    
    now = datetime.now().strftime("%H:%M:%S")
    sta = datetime.now()
    print(sta)
    
    # Start getting an Appointment
            
    # Change to Dubai City
    dubai()

    # Select Calendar Table
    try:
        # 10 is the maximum time to wait
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.ID, "appointments_consulate_appointment_date"))).click()
    except:
        print("Appointment dropdown is closed-1, Check for BAN")
        return
    
    # Search in Calendar for the Month    
    for n in range(30):
        try:
            # 10 is the maximum time to wait
            month = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[@id=\"ui-datepicker-div\"]/div[1]/div/div"))).text
        except:
            print("Reading Month failed-1")
            break

        if mon in month:
            break
        else:
            driver.find_element(By.CLASS_NAME, "ui-datepicker-next").click()
            
    # Select first availbale date 
    try:
        driver.find_element(By.XPATH, "(//td[@data-handler='selectDay'])[1]").click()
        print("Selecting date was successful")

    except selenium.common.exceptions.NoSuchElementException:
        print("Selecting date was unsuccessful")
        return
    
    # Select time of appointment
    try:
        # 10 is the maximum time to wait
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (By.XPATH,
                    "//*[@id=\"appointments_consulate_appointment_time\"]/option[2]")))
    except:
        print("No Time Available to select")

    try:
        Time = driver.find_element(By.ID, "appointments_consulate_appointment_time")
        Time_1 = Select(Time)
        Time_1.select_by_index(1)
        print("Selecting time was successful")
    except:
        print("Selecting time was unsuccessful")
        return
    
    # Click on reschedule button
    try:
        # 10 is the maximum time to wait
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//input[@id='appointments_submit']"))).click()

    except TimeoutException:
        print("click submit failed")
        return

    "////////////////////////////////////////////////////////////////////////////"
    # Cancel Button
    # Uncomment this section if you want to test the bot (ctrl + /)

    # try:
    #     # 10 is the maximum time to wait
    #     element = WebDriverWait(driver, 10).until(
    #         EC.element_to_be_clickable(
    #             (By.XPATH, "//a[normalize-space()='Cancel']"))).click()
    # except TimeoutException:
    #     pass
    "////////////////////////////////////////////////////////////////////////////"
    
    # Click on Confirm Button
    # comment this section if you want to test the bot (ctrl + /)
    
    try:
        # 10 is the maximum time to wait
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[normalize-space()='Confirm']"))).click()

    except TimeoutException:
        return

    # Calculate total duration
    sto = datetime.now()
    du = (sto - sta).total_seconds()
    print(" Duration : " + str(du))

    sleep(3)
    
    # check if reschedule was successful or not
    c_url = driver.current_url

    if "instructions" in c_url:
        alert = driver.find_element(By.CSS_SELECTOR, ".notice").text
        if "You Have Successfully Scheduled Your Appointment" in alert:
            print("Appointment time has been Rescheduled Successfully: " + user)
            driver.quit() 
        else:
            print("Confirm was un--successful")
            sleep(60)
            driver.get(link)
            refresh()
    else:
        # refresh appointment page
        driver.get(link)       
