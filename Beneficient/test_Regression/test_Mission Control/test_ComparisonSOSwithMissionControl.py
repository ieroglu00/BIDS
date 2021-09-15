import re
import time

from _pytest.outcomes import fail
from selenium.webdriver import ActionChains
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
@allure.description("Test case to verfiy CompareAmountSOSandMissionControl")
@allure.severity(severity_level="High")
def test_CompareAmountSOSandMissionControl(test_setup):
      DropDownCount=2
      ComparisonCounter=0
      YearList = []

      AmountListFundLevel = []
      AmountListInvestmentLevel = []

      #Lists for SOSFunds/Investmenst
      AmountListSOSFunds = []
      AmountListSOSInvestmenst = []

      driver.find_element_by_xpath("//*[@title='Quarterly NAV Close']").click()
      for ia in range(1000):
            try:
                  bool = driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
            except Exception:
                  #print("Loader finished")
                  time.sleep(3)
                  break
      #print(driver.title)
      assert "Quarterly NAV Close - BIDS" in driver.title, "Failed to open Quarterly NAV Close Page"
      driver.find_element_by_xpath("//strong[contains(text(),'Mission Control')]").click()
      for ia in range(1000):
            try:
                  bool = driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
            except Exception:
                  #print("Loader finished")
                  time.sleep(3)
                  break
      assert "COR_ReportMissionControl - BIDS" in driver.title, "Failed to open Mission Control page"

      for ia in range(DropDownCount):
            #print("Count is " + str(ia))
            if ia == 0:
                  # print("First Data")
                  P = driver.find_element_by_xpath(
                        "//div/span[@class='DropdownWidget---accessibilityhidden']").text
            elif ia > 0:
                  time.sleep(5)
                  # print("Other Data")
                  driver.find_element_by_xpath(
                        "//input[@class='PickerWidget---picker_input PickerWidget---placeholder']").send_keys()
                  elements = driver.find_elements_by_xpath(
                        "//div[@class='DropdownWidget---dropdown_value DropdownWidget---inSideBySideItem']")
                  for elem in elements:
                        elem.click()
                        break
                  time.sleep(5)
                  ActionChains(driver).key_down(Keys.DOWN).perform()
                  time.sleep(3)
                  ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
                  for ia in range(1000):
                        try:
                              bool = driver.find_element_by_xpath(
                                    "//div[@id='appian-working-indicator-hidden']").is_enabled()
                              # print("Loader present")
                        except Exception:
                              # print("Loader finished")
                              time.sleep(5)
                              break
                  P = driver.find_element_by_xpath(
                        "//div/span[@class='DropdownWidget---accessibilityhidden']").text
            for ia in range(1000):
                  try:
                        bool = driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
                  except Exception:
                        # print("Loader finished")
                        time.sleep(3)
                        break
            time.sleep(5)
            AmtFundLevel = driver.find_element_by_xpath(
                  "//div[@class='ContentLayout---content_layout']/div[5]/div[2]/div/div/div[2]/div[10]/div[2]/div/p/span/strong").text
            AmtInvestmentLevel = driver.find_element_by_xpath(
                  "//div[@class='ContentLayout---content_layout']/div[5]/div[2]/div/div/div[3]/div[10]/div[2]/div/p/span/strong").text

            YearList.append(P)
            AmountListFundLevel.append(AmtFundLevel)
            AmountListInvestmentLevel.append(AmtInvestmentLevel)


      #Navigating back to Quarterly NAV Close Page
      #print("Now Comparing the numbers for SOS Funds*******************")
      driver.find_element_by_xpath("//*[@title='Quarterly NAV Close']").click()
      for ia in range(1000):
            try:
                  bool = driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
            except Exception:
                  #print("Loader finished")
                  time.sleep(3)
                  break
      #print(driver.title)
      assert "Quarterly NAV Close - BIDS" in driver.title, "Failed to open Quarterly NAV Close Page"
      driver.find_element_by_xpath("//strong[contains(text(),'Sign-Off Summary: Funds')]").click()
      for ia in range(1000):
            try:
                  bool = driver.find_element_by_xpath(
                        "//div[@id='appian-working-indicator-hidden']").is_enabled()
                  # print("Loader present")
            except Exception:
                  # print("Loader finished")
                  time.sleep(5)
                  break
      for iaa in range(DropDownCount):
            # print("Count is " + str(ia))
            if iaa == 0:
                  # print("First Data")
                  P = driver.find_element_by_xpath(
                        "//div[@class='ContentLayout---content_layout']/div[1]/div/div[3]/div[1]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div").text
            elif iaa > 0:
                  time.sleep(5)
                  # print("Other Data")
                  driver.find_element_by_xpath(
                        "//input[@class='PickerWidget---picker_input PickerWidget---placeholder']").send_keys()
                  elements = driver.find_elements_by_xpath(
                        "//div[@class='ContentLayout---content_layout']/div[1]/div/div[3]/div[1]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div")
                  for elem in elements:
                        elem.click()
                        break
                  time.sleep(5)
                  ActionChains(driver).key_down(Keys.DOWN).perform()
                  time.sleep(3)
                  ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
                  for ia in range(1000):
                        try:
                              bool = driver.find_element_by_xpath(
                                    "//div[@id='appian-working-indicator-hidden']").is_enabled()
                              # print("Loader present")
                        except Exception:
                              # print("Loader finished")
                              time.sleep(5)
                              break
                  P = driver.find_element_by_xpath(
                        "//div[@class='ContentLayout---content_layout']/div[1]/div/div[3]/div[1]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div").text
            for ia in range(1000):
                  try:
                        bool = driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
                  except Exception:
                        # print("Loader finished")
                        time.sleep(3)
                        break
            time.sleep(5)
            element = driver.find_element_by_xpath("//div[@class='ContentLayout---content_layout']/div[3]/div/div/div[2]/div/div[8]/div/div[2]/div/input")
            FundsSOSTotalPartnerNAVEndingUSD = element.get_attribute("value")
            AmountListSOSFunds.append(FundsSOSTotalPartnerNAVEndingUSD)


             # Navigating back to Quarterly NAV Close Page

      driver.find_element_by_xpath("//*[@title='Quarterly NAV Close']").click()
      for ia in range(1000):
            try:
                  bool = driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
            except Exception:
                  # print("Loader finished")
                  time.sleep(3)
                  break
      # print(driver.title)
      assert "Quarterly NAV Close - BIDS" in driver.title, "Failed to open Quarterly NAV Close Page"
      driver.find_element_by_xpath("//strong[contains(text(),'Sign-Off Summary: Investments')]").click()
      for ia in range(1000):
            try:
                  bool = driver.find_element_by_xpath(
                        "//div[@id='appian-working-indicator-hidden']").is_enabled()
                  # print("Loader present")
            except Exception:
                  # print("Loader finished")
                  time.sleep(5)
                  break
      for iaaa in range(DropDownCount):
            # print("Count is " + str(ia))
            if iaaa == 0:
                  # print("First Data")
                  P = driver.find_element_by_xpath(
                        "//div[@class='ContentLayout---content_layout']/div[2]/div/div/div/div[2]/div[1]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div").text
            elif iaaa > 0:
                  time.sleep(5)
                  # print("Other Data")
                  driver.find_element_by_xpath(
                        "//input[@class='PickerWidget---picker_input PickerWidget---placeholder']").send_keys()
                  elements = driver.find_elements_by_xpath(
                        "//div[@class='ContentLayout---content_layout']/div[2]/div/div/div/div[2]/div[1]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div")
                  for elem in elements:
                        elem.click()
                        break
                  time.sleep(5)
                  ActionChains(driver).key_down(Keys.DOWN).perform()
                  time.sleep(3)
                  ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
                  for ia in range(1000):
                        try:
                              bool = driver.find_element_by_xpath(
                                    "//div[@id='appian-working-indicator-hidden']").is_enabled()
                              # print("Loader present")
                        except Exception:
                              # print("Loader finished")
                              time.sleep(5)
                              break
                  P = driver.find_element_by_xpath(
                        "//div[@class='ContentLayout---content_layout']/div[2]/div/div/div/div[2]/div[1]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div").text
            for ia in range(1000):
                  try:
                        bool = driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
                  except Exception:
                        # print("Loader finished")
                        time.sleep(3)
                        break
            time.sleep(5)
            element = driver.find_element_by_xpath(
                  "//div[@class='ContentLayout---content_layout']/div[5]/div/div/div[2]/div/div[8]/div/div[2]/div/input")
            InvestmentSOSBenLevelStruckNAVatEndofPeriod = element.get_attribute("value")
            AmountListSOSInvestmenst.append(InvestmentSOSBenLevelStruckNAVatEndofPeriod)

      #print("Now Showing the numbers ")
      # print("Date - " + "Fund Level - " + "Investment Level - " + "SOS Funds - " + "SOS Investment")
      # for x in range(len(AmountListFundLevel)):
      #      print(YearList[x]+" - "+AmountListFundLevel[x]+" - "+AmountListInvestmentLevel[x]+" - "+AmountListSOSFunds[x]+" - "+AmountListSOSInvestmenst[x])
      #      #print("AmountInvestmentLevel: "+AmountListInvestmentLevel[x])
      #      print()

      #print("Now Comparing the numbers")
      AmountListFundLevel.sort()
      AmountListInvestmentLevel.sort()
      AmountListSOSFunds.sort()
      AmountListSOSInvestmenst.sort()

      for x1 in range(len(AmountListFundLevel)):
            if AmountListFundLevel[x1] != AmountListSOSFunds[x1]:
                  print("AmountListFundLevel ( " + AmountListFundLevel[x1] + " ) and AmountListSOSFunds ( " +
                        AmountListSOSFunds[x1] + " ) **NOT** matching for year " + YearList[x1])
            # else:
            #       print("AmountListFundLevel ( " + AmountListFundLevel[x1] + " ) and AmountListSOSFunds ( " +
            #             AmountListSOSFunds[x1] + " ) matching for year " + YearList[x1])

      for x11 in range(len(AmountListFundLevel)):
            if AmountListInvestmentLevel[x11] != AmountListSOSInvestmenst[x11]:
                  ComparisonCounter = ComparisonCounter + 1
                  print("AmountListInvestmentLevel ( " + AmountListInvestmentLevel[
                        x11] + " ) and AmountListSOSInvestmenst ( " + AmountListSOSInvestmenst[
                              x11] + " ) **NOT** matching for year " + YearList[x11])
            # else:
            #       print("AmountListInvestmentLevel ( " + AmountListInvestmentLevel[
            #             x11] + " ) and AmountListSOSInvestmenst ( " + AmountListSOSInvestmenst[
            #                   x11] + " ) matching for year " + YearList[x11])

      if (ComparisonCounter>0):

            pytest.fail("Fund / Investment Level amount in Mission Control for different periods in not matching with SOSFund / SOSInvestment amount")
