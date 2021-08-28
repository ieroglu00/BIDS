import random
import time

from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException
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
  driver.get("https://beneficientdev.appiancloud.com/suite/")
  enter_username("neeraj.kumar")
  enter_password("Motorola@408")
  driver.find_element_by_xpath("//input[@type='submit']").click()

  # yield
  # driver.quit()

@pytest.mark.regression
@allure.description("Test case to verify Dropdown")
@allure.severity(severity_level="High")
def test_dropdown(test_setup):
    #time.sleep(7)
    # new = driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
    # WebDriverWait(driver, timeout=25).until(new)

    driver.find_element_by_xpath("//*[@title='Quarterly NAV Close']").click()

    # driver.find_element_by_xpath("//*[@title='Quarterly NAV Close']").click()
    for ia in range(1000):
         try:
           bool=driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
         except Exception:
           print("Loader finished")
           break
    driver.find_element_by_xpath("//strong[contains(text(),'Mission Control')]").click()
    for ia in range(1000):
         try:
           bool=driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
         except Exception:
           print("Loader finished")
           break





