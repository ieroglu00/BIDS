import time
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
  driver=webdriver.Chrome(executable_path="C:/BIDS/beneficienttest/Beneficient/Chrome/chromedriver.exe")
  driver.implicitly_wait(10)
  driver.maximize_window()
  driver.get("https://beneficientdev.appiancloud.com/suite/")
  enter_username("neeraj.kumar")
  enter_password("Crochet@786")
  driver.find_element_by_xpath("//input[@type='submit']").click()

  yield
  driver.quit()

@pytest.mark.regression
@allure.description("Test case to verify Delete error in Detailed Cap Stmts inside Beneficient")
@allure.severity(severity_level="High")
def test_NavToDetailedCAPSTMTs(test_setup):
  try:
    driver.find_element_by_xpath("//tbody/tr[1][@class='PagingGridLayout---selectable']/td[2]").click()
    print("Fund at Beneficient listing is clicked")
  except Exception:
    print("Fund at Beneficient listing is not clickable")
    allure.attach(driver.get_screenshot_as_png(), name="Image1", attachment_type=allure.attachment_type.PNG)
    pytest.fail("Failed to load List of Beneficient")

  time.sleep(5)
  try:
    driver.find_element_by_xpath("//button[text()='Detailed Cap Stmts']").click()
    print("Detailed Cap Stmts tab is clicked")
  except Exception:
    print("Detailed Cap Stmts tab is not clickable")
    allure.attach(driver.get_screenshot_as_png(), name="Image1", attachment_type=allure.attachment_type.PNG)
    pytest.fail("Failed to click on Detailed Cap Stmts tab")

  time.sleep(5)


  Check=0
  try:
   if driver.find_element_by_xpath("//a[contains(text(),'Add/Edit Detailed Cap Statement')]").is_displayed():
    Check=1
  except Exception:
   if driver.find_element_by_xpath("//button[text()='Cancel']").is_displayed():
    Check=2

  if   Check==1:
   driver.find_element_by_xpath("//a[contains(text(),'Add/Edit Detailed Cap Statement')]").click()
  elif Check==2:
   driver.find_element_by_xpath("//button[text()='Cancel']").click()
   driver.find_element_by_xpath("//a[contains(text(),'Add/Edit Detailed Cap Statement')]").click()

  time.sleep(3)
  driver.find_element_by_xpath("//td[2][@class='EditableGridLayout---reducedPadding']/div/input[1]").clear()

  driver.find_element_by_xpath("//td[2][@class='EditableGridLayout---reducedPadding']/div/input[1]").send_keys("12")
  time.sleep(3)
  Data = driver.find_element_by_xpath("//td[2][@class='EditableGridLayout---reducedPadding']/div/input[1]").text
  time.sleep(5)
  driver.find_element_by_xpath("//button[text()='Save']").click()
  driver.find_element_by_xpath("//a[contains(text(),'Add/Edit Detailed Cap Statement')]").click()
  driver.find_element_by_xpath("//p/a[text()='Delete']").click()

  driver.find_element_by_xpath("//button[text()='Save']").click()
  time.sleep(5)
  Data=driver.find_element_by_xpath("//td[2]/p[@class='ParagraphText---richtext_paragraph ParagraphText---default_direction ParagraphText---align_end elements---global_p']").text

  if Data in "12":
    allure.attach(driver.get_screenshot_as_png(), name="Image1", attachment_type=allure.attachment_type.PNG)
    pytest.fail("Delete functionality is not working as expected")

