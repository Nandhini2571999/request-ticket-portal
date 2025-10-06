from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager



class MohreBot:
    def __init__(self, headless=False):
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument("--headless")

        service = Service(ChromeDriverManager().install())  
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.implicitly_wait(10)

    def go_to_url(self, url):
        self.driver.get(url)
        self.driver.find_element(By.ID, 'alertCloseEst').click()
        self.driver.find_element(By.ID, 'iAccept').click()
        self.driver.find_element(By.XPATH, '//*[@id="LoginMethodForm"]/a').click()
        windows = self.driver.window_handles
        self.driver.switch_to.window(windows[1])
        self.driver.find_element(By.XPATH, '//*[@id="username"]').send_keys('0547711039')
        self.driver.find_element(By.ID, 'basicPasswordForm-submitButton').click()
        time.sleep(180)
        print("Successfully Login!")


obj = MohreBot()
obj.go_to_url('https://eservices.mohre.gov.ae/tasheelweb/account/login')
obj.login()
print("Yes")