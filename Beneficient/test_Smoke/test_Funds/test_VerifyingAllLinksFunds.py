import time

import xlrd
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
      driver.implicitly_wait(10)
      driver.maximize_window()
      driver.get("https://beneficienttest.appiancloud.com/suite/")
      enter_username("neeraj.kumar")
      enter_password("Crochet@786")
      driver.find_element_by_xpath("//input[@type='submit']").click()

      yield
      driver.quit()

@pytest.mark.regression
@allure.description("Test case to verfiy all links at Funds Page")
@allure.severity(severity_level="High")
def test_VerfyAllLinksFundsPage(test_setup):
    PageName = "Funds"
    PageTitle = "Funds - BIDS"
    loc = ("C:/Users/Neeraj/Desktop/New/Main.xls")

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    driver.find_element_by_xpath("//*[@title='" + PageName + "']").click()
    for iat5 in range(1000):
        try:
            bool = driver.find_element_by_xpath(
                "//div[@id='appian-working-indicator-hidden']").is_enabled()
        except Exception:
            time.sleep(1)
            break
    time.sleep(2)
    assert PageTitle in driver.title
    for ia in range(50):
        ia = ia + 1
        # print()
        # print("ia is " + str(ia))
        try:
            bool_series = pd.isnull(sheet.cell_value(ia, 0))
            # print("bool_series is "+ sheet.cell_value(ia, 0))
            if (bool_series == True):
                break
            else:
                if (sheet.cell_value(ia, 3) == "No"):
                    if (sheet.cell_value(ia, 0) == PageName):
                        print()
                        try:
                            InOrOut = sheet.cell_value(ia, 9)
                            # print("InOrOut is " + InOrOut)
                            if InOrOut == "Inside":
                                driver.find_element_by_xpath(sheet.cell_value(ia, 10)).click()
                                print("Parent Page link clicked ")
                                for iat2 in range(1000):
                                    try:
                                        bool = driver.find_element_by_xpath(
                                            "//div[@id='appian-working-indicator-hidden']").is_enabled()
                                    except Exception:
                                        time.sleep(1)
                                        break
                                time.sleep(3)
                                driver.find_element_by_xpath(sheet.cell_value(ia, 2)).click()
                            elif InOrOut == "Outside":
                                driver.find_element_by_xpath(sheet.cell_value(ia, 2)).click()

                            print("Verification started for:  " + sheet.cell_value(ia, 1))
                            for iat2 in range(1000):
                                try:
                                    bool = driver.find_element_by_xpath(
                                        "//div[@id='appian-working-indicator-hidden']").is_enabled()
                                except Exception:
                                    time.sleep(1)
                                    break
                            # print("link clicked:  " + sheet.cell_value(ia, 1))
                            # print("Skip is " + sheet.cell_value(ia, 3))
                            DoubleClick = sheet.cell_value(ia, 4)
                            # print("DoubleClick is "+DoubleClick)
                            NaviBack = sheet.cell_value(ia, 5)
                            # print("NaviBack is " + NaviBack)
                            TitleVerify = sheet.cell_value(ia, 6)
                            # print("TitleVerify is " + TitleVerify)
                            TitleToVerify = sheet.cell_value(ia, 7)
                            # print("TitleToVerify is " + TitleToVerify)
                            TitleLink = sheet.cell_value(ia, 8)
                            # print("TitleLink is " + TitleLink)

                            if DoubleClick == "Yes":
                                driver.find_element_by_xpath(sheet.cell_value(ia, 2)).click()
                                time.sleep(1)
                                # print("Link again clicked for  " + sheet.cell_value(ia, 1))

                            elif DoubleClick == "No":
                                # print("Inside Double clicked NO")
                                if NaviBack == "Yes" and TitleVerify == "No":
                                    # print("Inside NaviBack=Yes TitleVerify= NO")
                                    for iat3 in range(1000):
                                        try:
                                            bool = driver.find_element_by_xpath(
                                                "//div[@id='appian-working-indicator-hidden']").is_enabled()
                                        except Exception:
                                            time.sleep(2)
                                            break
                                    # print("Browser Back clicked for  " + sheet.cell_value(ia, 1))
                                    time.sleep(3)
                                    try:
                                        driver.find_element_by_xpath("//*[@title='" + PageName + "']").click()
                                    except Exception as e2:
                                        print(e2)
                                        driver.back()

                                    time.sleep(3)
                                elif NaviBack == "Yes" and TitleVerify == "Yes":
                                    # print("Inside NaviBack=Yes TitleVerify= Yes")
                                    for iat6 in range(1000):
                                        try:
                                            bool = driver.find_element_by_xpath(
                                                "//div[@id='appian-working-indicator-hidden']").is_enabled()
                                        except Exception:
                                            time.sleep(1)
                                            break
                                    TitleFound = driver.find_element_by_xpath(TitleLink).text
                                    # print("TitleFound is " + TitleFound)
                                    if (TitleFound != TitleToVerify):
                                        print("Something wrong found for  " + sheet.cell_value(ia, 1))
                                        print("TitleFound is " + TitleFound)
                                        print("Expected Title is " + TitleToVerify)
                                    else:
                                        print("Title matched for  " + sheet.cell_value(ia, 1))

                                    time.sleep(1)
                                    try:
                                        driver.find_element_by_xpath("//*[@title='" + PageName + "']").click()
                                        time.sleep(1)
                                        try:
                                            driver.switch_to_alert().accept()
                                        except Exception:
                                            pass
                                        for iat8 in range(1000):
                                            try:
                                                bool = driver.find_element_by_xpath(
                                                    "//div[@id='appian-working-indicator-hidden']").is_enabled()
                                            except Exception:
                                                time.sleep(1)
                                                break
                                        # print("Browser Back clicked 1")
                                    except Exception as e2:
                                        print(e2)
                                        driver.back()
                                        # print("Browser Back clicked 2")

                                elif NaviBack == "No" and TitleVerify == "Yes":
                                    # print("Inside NavBack no and Title Yes")
                                    for iat7 in range(1000):
                                        try:
                                            bool = driver.find_element_by_xpath(
                                                "//div[@id='appian-working-indicator-hidden']").is_enabled()
                                        except Exception:
                                            time.sleep(1)
                                            break
                                    TitleFound = driver.find_element_by_xpath(TitleLink).text
                                    # print("TitleFound1 is " + TitleFound)
                                    if (TitleFound1 != TitleToVerify):
                                        print("Something wrong found for  " + sheet.cell_value(ia, 1))
                                        print("TitleFound is " + TitleFound)
                                        print("Expected Title is " + TitleToVerify)
                                    else:
                                        print("Title matched for  " + sheet.cell_value(ia, 1))

                        except Exception as e:
                            print("Link not clicked / opened for  " + sheet.cell_value(ia, 1))
                            print(e)
                        for iat4 in range(1000):
                            try:
                                bool = driver.find_element_by_xpath(
                                    "//div[@id='appian-working-indicator-hidden']").is_enabled()
                            except Exception:
                                time.sleep(1)
                                break
        except Exception as e1:
            break
            print(e1)