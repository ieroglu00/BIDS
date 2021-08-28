import re
import time

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
@allure.description("Test case to verfiy difference in NAV Rollover from Prior Period")
@allure.severity(severity_level="High")
def test_differenceNAVRollOverPriorPeriod(test_setup):
      DropDownCount=6
      Differencecount = 0
      driver.find_element_by_xpath("//*[@title='Quarterly NAV Close']").click()
      for ia in range(1000):
            try:
                  bool = driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
            except Exception:
                  time.sleep(3)
                  break
      #print(driver.title)
      assert "Quarterly NAV Close - BIDS" in driver.title, "Failed to open Quarterly NAV Close Page"
      driver.find_element_by_xpath("//strong[contains(text(),'Mission Control')]").click()
      for ia in range(1000):
            try:
                  bool = driver.find_element_by_xpath("//div[@id='appian-working-indicator-hidden']").is_enabled()
            except Exception:
                  time.sleep(3)
                  break
      assert "COR_ReportMissionControl - BIDS" in driver.title, "Failed to open Mission Control page"
      #print(driver.title)
      Items = driver.find_elements_by_xpath(
            "//div[@class='BoxLayout---box BoxLayout---margin_below_standard'][1]/div[2]/div/div/div[4]/div")
      if len(Items) == 0:
            assert len(Items) != 0, "No Item available in NAV Rollover from Prior Period section"
      elif len(Items)>0:
            for ia in range(DropDownCount):
                  print("Count is " + str(ia))
                  if ia == 0:
                        #print("First Data")
                        P = driver.find_element_by_xpath(
                              "//div/span[@class='DropdownWidget---accessibilityhidden']").text
                  elif ia > 0:
                        time.sleep(5)
                        #print("Other Data")
                        driver.find_element_by_xpath("//input[@class='PickerWidget---picker_input PickerWidget---placeholder']").send_keys()
                        elements = driver.find_elements_by_xpath("//div[@class='DropdownWidget---dropdown_value DropdownWidget---inSideBySideItem']")
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
                                #print("Loader present")
                            except Exception:
                                #print("Loader finished")
                                time.sleep(5)
                                break
                        P = driver.find_element_by_xpath(
                              "//div/span[@class='DropdownWidget---accessibilityhidden']").text

                  time.sleep(7)
                  rows = driver.find_elements_by_xpath(
                        "//div[@class='BoxLayout---box BoxLayout---margin_below_standard'][1]/div[2]/div/div/div[4]/div")
                  for i in range(len(rows) + 1):
                        if i > 1:
                              DifferenceValueString = driver.find_element_by_xpath(
                                    "//div[@class='BoxLayout---box BoxLayout---margin_below_standard'][1]/div[2]/div/div/div[4]/div[" + str(
                                          i) + "]/div[2]/div/p").text
                              DifferenceValueString = DifferenceValueString.replace(" ", "")
                              DifferenceValueString = re.sub(r'[?|$|.|!|,|-]', r'', DifferenceValueString)
                              DifferenceValueString = ''.join(char for char in DifferenceValueString if char.isalnum())
                              if (len(DifferenceValueString) != 0):

                                    if (int(DifferenceValueString)>0):
                                          print("Difference found in [ "+P+" ] and the difference amount is "+DifferenceValueString)
                                          Differencecount=Differencecount+1



                  print()
      if (Differencecount>0):
            pytest.fail("Difference found in NAV Rollover from Prior Period")
      else:
           print("No Difference found in NAV Rollover from Prior Period")