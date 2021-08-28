import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pytest
from selenium import webdriver
import allure

@allure.step("Entering username ")
def enter_username(username):
  driver.find_element_by_id("un").send_keys(username)

@allure.step("Entering password ")
def enter_password(password):
  driver.find_element_by_id("pw").send_keys(password)

@pytest.fixture()
def test_setup():
  global driver
  driver=webdriver.Chrome(executable_path="C:\Laptop Data\Work\Python\chromedriver_win32 (1)\chromedriver")
  driver.implicitly_wait(10)
  driver.maximize_window()
  driver.get("https://beneficienttest.appiancloud.com/suite/")
  enter_username("neeraj.kumar")
  enter_password("Crochet@786")
  driver.find_element_by_xpath("//input[@type='submit']").click()

  yield
  driver.quit()

@pytest.mark.smoke
@allure.description("This is smoke testcase to verify all modules are opening")
@allure.severity(severity_level="High")
def test_ModulesVerify(test_setup):
    PageTitle1=driver.title
    time.sleep(5)
    print(PageTitle1+" Opened successfully")

    driver.find_element_by_xpath("//*[@title='Investments']").click()
    time.sleep(7)
    PageTitle2 = driver.title
    print(PageTitle2+" Opened successfully")

    driver.find_element_by_xpath("//*[@title='Transactions']").click()
    time.sleep(5)
    PageTitle3 = driver.title
    print(PageTitle3+" Opened successfully")

    driver.find_element_by_xpath("//*[@title='Liquid Trusts']").click()
    time.sleep(5)
    PageTitle4 = driver.title
    print(PageTitle4+" Opened successfully")

    driver.find_element_by_xpath("//*[@title='Quarterly NAV Closeh']").click()
    time.sleep(5)
    PageTitle5 = driver.title
    print(PageTitle5+" Opened successfully")



