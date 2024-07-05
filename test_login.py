import os

import pytest
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from orangehrm.Test_Excel_Functions.excel_functions import Selva_Excel_Functions
from orangehrm.Test_locators.locators import TestLocators


class Test_OrangeHRM:
    @pytest.fixture
    def boot(self):
        self.driver = webdriver.Chrome()
        self.driver.implicitly_wait(10)
        excel_file = 'C:\\Users\ssekar588\\PycharmProjects\\GUVI Selenium 2\\orangehrm\\Test_Datas\\test_data.xlsx'
        sheet_name = 'Sheet1'
        self.s = Selva_Excel_Functions(excel_file, sheet_name)
        self.rows = self.s.Row_Count()
        yield
        self.driver.close()

    def test_login(self, boot):
        self.driver.get(TestLocators.url)
        self.driver.maximize_window()
        wait = WebDriverWait(self.driver, 8)
        start_row = 2
        end_row = 5

        for row_no in range(start_row, end_row + 1):
            username = self.s.Read_Data(row_no, 6)
            password = self.s.Read_Data(row_no, 7)

            username_element = wait.until(EC.visibility_of_element_located((By.NAME, TestLocators().Email)))
            username_element.send_keys(username)

            password_element = wait.until(EC.visibility_of_element_located((By.NAME, TestLocators().Password)))
            password_element.send_keys(password)

            login_button = wait.until(EC.element_to_be_clickable((By.XPATH, TestLocators().Login_button)))
            login_button.click()

            try:
                wait.until(EC.url_matches('https://opensource-demo.orangehrmlive.com/web/index.php/dashboard/index'))
                self.s.Write_Data(row_no, 8, "TEST PASS")
                print("SUCCESS : Logged in with Username {a} & {b}".format(a=username, b=password))

                profile_image = wait.until(EC.presence_of_element_located((By.XPATH, TestLocators().Profile_image)))
                profile_image.click()
                logout_button = wait.until(EC.element_to_be_clickable((By.XPATH, TestLocators().Logout_button)))
                logout_button.click()

                wait.until(EC.url_matches("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login"))

            except TimeoutException:
                self.s.Write_Data(row_no, 8, "TEST FAIL")

                assert self.driver.current_url == 'https://opensource-demo.orangehrmlive.com/web/index.php/auth/login'
                print("FAIL : Login failed with Username {a} & {b}".format(a=username, b=password))
                screenshot_dir = os.path.join(os.getcwd(), 'screenshot')
                os.makedirs(screenshot_dir, exist_ok=True)
                screenshot_path = os.path.join(screenshot_dir, f"login_failure.png")
                self.driver.save_screenshot(screenshot_path)
                print(f"screenshot saved at: {screenshot_path}")
                self.driver.refresh()
