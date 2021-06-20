
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time, os, openpyxl

SIGNING_IN_URL = 'https://www.linkedin.com/'
NETWORK_PAGE = "http://linkedin.com/mynetwork"
WAIT_TIME = 5   # customize your wait time here
TIME_BETWEEN_CONNECTION = 2

class LinkedIn():   
 
    def __init__(self,username,password,chromeDriver_path):
        self.username = username
        self.password = password
        self.chromeDriver_path = chromeDriver_path
        self.driver = webdriver.Chrome(self.chromeDriver_path)

    def login(self):
        self.driver.get(SIGNING_IN_URL)
        #enter username and password
        username = self.driver.find_element_by_css_selector("#session_key")
        password = self.driver.find_element_by_css_selector("#session_password")
        username.send_keys(self.username)           # simulate typing into the element aka auto typing
        password.send_keys(self.password)
        # click the "sign in" button
        self.driver.find_element_by_css_selector("#main-content > section.section.section--hero > div.sign-in-form-container > form > button").click()


    def scroll_down_messaging(self):
        scroll_down_but = self.driver.find_elements_by_class_name('msg-overlay-bubble-header__control')[1]
        scroll_down_but.click()

    def quit_driver(self):
        self.driver.quit()

    def get_scroll_height(self):
        # get scroll height
        last_height = self.driver.execute_script("return document.body.scrollHeight")

        for i in range(3):
            # scroll down to bottom
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")

            # wait to load page
            time.sleep(WAIT_TIME)

            # calculate new scroll height and compare with last scroll height
            new_height = self.driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height


    def connect_people_myNetwork(self):

        count = 0

        self.driver.get(NETWORK_PAGE)
        # FIND PEOPLE ON "MY NETWORK":
        while True:
            try:
                count_limit = int(input("Please enter the number of people to auto connect: "))
            except ValueError:
                print("Invalid input. Please try again")
            else:
                break
        # navigate all connect buttons

        self.scroll_down_messaging()

        for _ in range(count_limit):
            try:
                WebDriverWait(self.driver,WAIT_TIME).until(EC.element_to_be_clickable((By.XPATH,"//span[text()='Connect']"))).click()
                count += 1
                time.sleep(TIME_BETWEEN_CONNECTION)
            except:
                print("Got intercepted")    # intercepted because of messaging bubble => try bs4

        print(str(count) + ' connection request sent')
