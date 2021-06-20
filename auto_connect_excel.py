
from auto_connect import LinkedIn
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time, os, openpyxl

WAIT_TIME = 5   # customize your wait time here
TIME_BETWEEN_CONNECTION = 2

class LinkedInExcel(LinkedIn):

    def __init__(self, username, password, chromeDriver_path, excel_file_path, excel_file_name,sheet_name):
        super().__init__(username, password, chromeDriver_path)

        self.excel_file_path = excel_file_path
        self.excel_file_name = excel_file_name
        self.sheet_name = sheet_name
    def get_excel_data(self):
    
        # change current dir
        os.chdir(self.excel_file_path)
        # open excel fle
        tracking_file = openpyxl.load_workbook(self.excel_file_name, read_only=True)
        client_sheet = tracking_file.get_sheet_by_name(self.sheet_name)    # customized the sheet here
        return client_sheet

    def analyze_client_sheet(self,client_sheet,client_range,name_col,url_col):
        # info with keys are names and values are url
        client_info = {}    

        # customized rows and columns (this depends on the file)

        for i in range(2, client_range + 1 ):    
            client_name = client_sheet.cell(i, name_col).value     # i is row and 4 is column number
            client_url = client_sheet.cell(i,url_col).value
            client_info[client_name] = client_url

        return client_info

    def connect_people_excelFile(self, list_client):
        connect_url_list = list(list_client.values())     
        client_name = list(list_client.keys())


        for i in range(len(connect_url_list)):
            self.driver.get(connect_url_list[i])

            # navigate "connect" and "add note" buttons
            connect_but = WebDriverWait(self.driver,WAIT_TIME).until(lambda x:x.find_element_by_class_name('pvs-profile-actions__action'))
            connect_but.click()
            try:
                self.driver.find_element_by_class_name('mr1').click()
            # move to the next client if we cannot message the previous client (possibly due to LinkedIn Premium)
            except: 
                print("No such element. Moving on to the next client")
                continue
            else:
                # what if clients have premium? => if 
                # send messages with name and info
                customMessage = "Hi " + str(client_name[i]) + " customized_message "
                elementID = self.driver.find_element_by_class_name('msg-form__contenteditable')
                elementID.send_keys(customMessage)
                WebDriverWait(self.driver,WAIT_TIME).until(EC.element_to_be_clickable((By.XPATH,"//span[text()='Send']"))).click()

                # count how many people have sent
                count += 1 
                self.driver.back()
                time.sleep(TIME_BETWEEN_CONNECTION)
                
                if count == len(connect_url_list):
                    print(str(count) + ' connection request sent')