import time

import xlrd
from openpyxl import Workbook
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.expected_conditions import staleness_of
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pytest
from selenium import webdriver
import allure
import pandas as pd

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
      # driver.implicitly_wait(10)
      # driver.maximize_window()
      # driver.get("https://beneficienttest.appiancloud.com/suite/")
      # enter_username("neeraj.kumar")
      # enter_password("Crochet@786")
      # driver.find_element_by_xpath("//input[@type='submit']").click()

      yield
      driver.quit()

@pytest.mark.regression
@allure.description("Test case to verfiy all links at Funds Inside Page")
@allure.severity(severity_level="High")
def test_VerfyAllLinksFundsInsidePage(test_setup):
    loc = ('C:/Users/Neeraj/PycharmProjects/beneficienttest/Beneficient/Testt/TestExcel.xls')
    wb1 = Workbook()
    ws = wb1.active
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    print()
    for i in range(0, 4):
           print(sheet.cell_value(i, 0))

           if sheet.cell_value(i, 0).strip():
             print("Not Empty")

           else:
             print(i)
             if (i==0):
                 i=i+1
             print("Empty")
             ws.cell(row=i, column=1).value = "Test"
             wb1.save("TestExcel.xls")
             if (i==1):
                 i=i-1

    for i1 in range(0, 4):
           print(sheet.cell_value(i1, 0))




