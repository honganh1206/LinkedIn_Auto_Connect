
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time, os, openpyxl

SIGNING_IN_URL = 'https://www.linkedin.com/'
NETWORK_PAGE = "http://linkedin.com/mynetwork"
WAIT_TIME = 5   # customize your wait time here


class LinkedIn():   
 
    def __init__(self,username,password,chromeDriver_path):
        self.username = username
        self.password = password
        self.chromeDriver_path = chromeDriver_path
        self.excel_file_path = "excel_file_path"
        self.excel_file_name = 'file_name.xlsx'

    def set_up_driver(self):
        self.driver = webdriver.Chrome(self.chromeDriver_path)
        driver = self.driver
        return driver

    def login(self, driver):
        driver.get(SIGNING_IN_URL)
        #enter username and password
        username = driver.find_element_by_css_selector("#session_key")
        password = driver.find_element_by_css_selector("#session_password")
        username.send_keys(self.username)           # simulate typing into the element aka auto typing
        password.send_keys(self.password)
        time.sleep(2)
        # click the "sign in" button
        sign_in_but = WebDriverWait(driver,WAIT_TIME).until(lambda x:x.find_element_by_css_selector("#main-content > section.section.section--hero > div.sign-in-form-container > form > button"))
        sign_in_but.click()


    def scroll_down_messaging(self, driver):
        scroll_down_but = driver.find_elements_by_class_name('msg-overlay-bubble-header__control')[1]
        scroll_down_but.click()

    def quit_driver(self,driver):
        self.driver.quit()

    def get_scroll_height(self,driver):
        # get scroll height
        last_height = driver.execute_script("return document.body.scrollHeight")

        for i in range(3):
            # scroll down to bottom
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")

            # wait to load page
            time.sleep(WAIT_TIME)

            # calculate new scroll height and compare with last scroll height
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

    
    def connect_people_myNetwork(self,driver):

        count = 0

        driver.get(NETWORK_PAGE)
            # FIND PEOPLE ON "MY NETWORK":
        while True:
            try:
                count_limit = int(input("Please enter the number of people to auto connect: "))
            except ValueError:
                print("Invalid input. Please try again")
            else:
                break
        # navigate all connect buttons

        self.scroll_down_messaging(driver)

        for _ in range(count_limit):
            try:
                WebDriverWait(driver,WAIT_TIME).until(EC.element_to_be_clickable((By.XPATH,"//span[text()='Connect']"))).click()
                count += 1
                time.sleep(2)
            except:
                print("Got intercepted")    # intercepted because of messaging bubble => try bs4

        print(str(count) + ' connection request sent')

    def get_excel_data(self,driver):
        
        # change current dir
        os.chdir(self.excel_file_path)
        # open excel fle
        tracking_file = openpyxl.load_workbook(self.excel_file_name)
        client_sheet = tracking_file.get_sheet_by_name('sheet_name')    # customized the sheet here

        # info with keys are names and values are url
        client_info = {}    

        # customized rows and columns (this depends on the file)

        for i in range(2, 71):    
            client_name = client_sheet.cell(i, 4).value     # i is row and 4 is column number
            client_url = client_sheet.cell(i,6).value
            client_info[client_name] = client_url

        return client_info
        
    def connect_people_excelFile(self,driver, list_client):
        connect_url_list = list(list_client.values())       # convert to a list because dict.values() are not subscriptable
        client_name = list(list_client.keys())


        for i in range(len(connect_url_list)):
            driver.get(connect_url_list[i])

            # navigate "connect" and "add note" buttons
            connect_but = WebDriverWait(driver,10).until(lambda x:x.find_element_by_class_name('pvs-profile-actions__action'))
            connect_but.click()
            driver.find_element_by_class_name('mr1').click()
            
            # send messages with name and info
            customMessage = "Hi " + str(client_name[i]) + " customized_message "
            elementID = driver.find_element_by_id('custom-message')
            elementID.send_keys(customMessage)
            time.sleep(WAIT_TIME)
            WebDriverWait(driver,WAIT_TIME).until(EC.element_to_be_clickable((By.XPATH,"//span[text()='Send']"))).click()

            # count how many people have sent
            count += 1 
            driver.back()
            time.sleep(2)
            
            if count == len(connect_url_list):
                print(str(count) + ' connection request sent')

                    
