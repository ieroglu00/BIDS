from datetime import datetime, timedelta,date
import math
import re
import time
import openpyxl
from fpdf import FPDF
import pytest
from selenium import webdriver
import allure
import imaplib
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pyodbc
from selenium.webdriver.chrome.options import Options

#----------------SalesForce Username and Password IDs-------------
@allure.step("Entering username ")
def enter_username(username):
  driver.find_element_by_id("username").send_keys(username)

@allure.step("Entering password ")
def enter_password(password):
  driver.find_element_by_id("password").send_keys(password)

@pytest.fixture()
def test_setup():
  global driver
  global TestName
  global description
  global TestResult
  global TestResultStatus
  global TestDirectoryName
  global path
  global FundNameList
  global FundNameListAfterRemove
  global ct
  global Exe
  global D1
  global D2
  global d1
  global d2
  global DollarDate
  global FundToOpen
  global TotalFundsLengh
  global FieldDataBIDS
  global FieldDataSF

  TestName = "test_DealLog_SFBIDSPhase1"
  description = "This test scenario is to verify integration between Sales Force and BIDS application by creating an opportunity"
  TestResult = []
  TestResultStatus = []
  TestFailStatus = []
  FailStatus="Pass"
  TestDirectoryName = "test_DealLog(SF-BIDS)"
  Exe="Yes"
  Directory = 'test_DealLog(SF-BIDS)/'
  path = 'C:/BIDS/beneficienttest/Beneficient/' + Directory

  FundNameList=[]
  FundNameListAfterRemove=[]
  FieldDataBIDS = {}
  FieldDataSF = {}

  ExcelFileName = "Execution"
  locx = (path+'Executiondir/' + ExcelFileName + '.xlsx')
  wbx = openpyxl.load_workbook(locx)
  sheetx = wbx.active

  for ix in range(1, 100):
      if sheetx.cell(ix, 1).value == None:
          break
      else:
          if sheetx.cell(ix, 1).value == TestName:
              if sheetx.cell(ix, 2).value == "No":
                  Exe="No"
              elif sheetx.cell(ix, 2).value == "Yes":
                  Exe="Yes"

  if Exe=="Yes":
      #-----------Disabling access popup from Chrome------------------
      option = Options()
      option.add_argument("--disable-infobars")
      option.add_argument("start-maximized")
      option.add_argument("--disable-extensions")
      option.add_experimental_option("prefs", {"profile.default_content_setting_values.notifications": 2})

      driver=webdriver.Chrome(chrome_options=option, executable_path="C:/BIDS/beneficienttest/Beneficient/Chrome/chromedriver.exe")
      driver.implicitly_wait(10)
      driver.maximize_window()

      #---------------------For Login in SalesForce--------------
      try:
          driver.get("https://beneficient--int.my.salesforce.com/")
          enter_username("neeraj.kumar@bitsinglass.com.int")
          enter_password("Crochet@786")
          driver.find_element_by_id("Login").click()
          time.sleep(10)
          TestResult.append(
              " Sales Force site launched successfully")
          TestResultStatus.append("Pass")
      except Exception:
        PageLoadError=driver.find_element_by_xpath("//span[@jsselect='heading']").text
        print(PageLoadError)
        TestResult.append(
            " Sales Force site is not able to load. Below error found\n" + PageLoadError)
        TestResultStatus.append("Fail")
        driver.quit()

      ct = datetime.now().strftime("%d_%B_%Y_%I_%M%p")
      ctReportHeader = datetime.now().strftime("%d %B %Y %I %M%p")

      today = date.today()
      D1=today.strftime("%Y-%m-%d")
      d1=D1
      DollarDate=datetime.strptime(d1, '%Y-%m-%d')
      DollarDate="$"+DollarDate.date().__str__()+"$"
      d1 = datetime.strptime(D1, "%Y-%m-%d")

  yield
  if Exe == "Yes":
      class PDF(FPDF):
          def header(self):
              self.image(path+'EmailReportContent/Ben.png', 10, 8, 33)
              self.set_font('Arial', 'B', 15)
              self.cell(73)
              self.set_text_color(0, 0, 0)
              self.cell(35, 10, ' Test Report ', 1, 1, 'B')
              self.set_font('Arial', 'I', 10)
              self.cell(150)
              self.cell(30, 10, ctReportHeader, 0, 0, 'C')
              self.ln(20)

          def footer(self):
              self.set_y(-15)
              self.set_font('Arial', 'I', 8)
              self.set_text_color(0, 0, 0)
              self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')

      pdf = PDF()
      pdf.alias_nb_pages()
      pdf.add_page()
      pdf.set_font('Times', '', 12)
      pdf.cell(0, 10, "Test Case Name:  "+TestName, 0, 1)
      pdf.multi_cell(0, 10, "Description:  "+description, 0, 1)

      for i1 in range(len(TestResult)):
         pdf.set_fill_color(255, 255, 255)
         pdf.set_text_color(0, 0, 0)
         if (TestResultStatus[i1] == "Fail"):
             #print("Fill Red color")
             pdf.set_text_color(255, 0, 0)
             TestFailStatus.append("Fail")
         TestName1 = TestResult[i1].encode('latin-1', 'ignore').decode('latin-1')
         pdf.multi_cell(0, 7,str(i1+1)+")  "+TestName1, 0, 1,fill=True)
         TestFailStatus.append("Pass")
      pdf.output(TestName+"_" + ct + ".pdf", 'F')

      #-----------To check if any failed Test case present-------------------
      for io in range(len(TestResult)):
          if TestFailStatus[io]=="Fail":
              FailStatus="Fail"
      # ---------------------------------------------------------------------

      # -----------To add test case details in PDF details sheet-------------
      ExcelFileName = "FileName"
      loc = (path+'PDFFileNameData/' + ExcelFileName + '.xlsx')
      wb = openpyxl.load_workbook(loc)
      sheet = wb.active
      print()
      check = TestName
      PdfName = TestName + "_" + ct + ".pdf"
      checkcount = 0

      for i in range(1, 100):
          if sheet.cell(i, 1).value == None:
              if checkcount == 0:
                  sheet.cell(row=i, column=1).value = check
                  sheet.cell(row=i, column=2).value = PdfName
                  sheet.cell(row=i, column=3).value = TestDirectoryName
                  sheet.cell(row=i, column=4).value = description
                  sheet.cell(row=i, column=5).value = FailStatus
                  checkcount = 1
              wb.save(loc)
              break
          else:
              if sheet.cell(i, 1).value == check:
                  if checkcount == 0:
                    sheet.cell(row=i, column=2).value = PdfName
                    sheet.cell(row=i, column=3).value = TestDirectoryName
                    sheet.cell(row=i, column=4).value = description
                    sheet.cell(row=i, column=5).value = FailStatus
                    checkcount = 1
      #----------------------------------------------------------------------------

      #---------------------To add Test name in Execution sheet--------------------
      ExcelFileName1 = "Execution"
      loc1 = (path+'Executiondir/' + ExcelFileName1 + '.xlsx')
      wb1 = openpyxl.load_workbook(loc1)
      sheet1 = wb1.active
      checkcount1 = 0

      for ii1 in range(1, 100):
          if sheet1.cell(ii1, 1).value == None:
              if checkcount1 == 0:
                  sheet1.cell(row=ii1, column=1).value = check
                  checkcount1 = 1
              wb1.save(loc1)
              break
          else:
              if sheet1.cell(ii1, 1).value == check:
                  if checkcount1 == 0:
                    sheet1.cell(row=ii1, column=1).value = check
                    checkcount1 = 1
      #-----------------------------------------------------------------------------

      #driver.quit()

@pytest.mark.smoke
def test_DealLog_SFBIDSPhase1(test_setup):
    if Exe == "Yes":
        SHORT_TIMEOUT = 5
        LONG_TIMEOUT = 400
        #LOADING_ELEMENT_XPATH = "//div[@id='appian-working-indicator-hidden']"
        LOADING_ELEMENT_XPATH = "//div[@class='slds-spinner_container slds-grid']"
        try:
            # ----------------------Reading Field data from reference sheet-----------------
            ExcelFileName = "FieldData"
            loc = (path + 'Reference Data/' + ExcelFileName + '.xlsx')
            wb = openpyxl.load_workbook(loc)
            sheet = wb.active
            for fielddata in range(1,50):
                Value1 = sheet.cell(row=fielddata, column=2).value
                try:
                    key1=sheet.cell(row=fielddata, column=1).value
                    FieldDataBIDS[key1] = Value1
                except Exception:
                    pass
                try:
                    key2 = sheet.cell(row=fielddata, column=4).value
                    FieldDataSF[key2] = Value1
                except Exception:
                    pass

            print(FieldDataBIDS)
            print(FieldDataSF)
            # print(FieldDataBIDS['Opportunity Owner'])
            # print(FieldDataSF['Opportunity.SubStage__c'])

            #------------------------Get verification code from Gmail---------------------------------
            host = 'imap.gmail.com'
            username = 'neeraj.kumar@bitsinglass.com'
            password = 'MotoCrochet@786'

            # -------------Function to get email content part i.e its body part
            def get_body(msg):
                if msg.is_multipart():
                    return get_body(msg.get_payload(0))
                else:
                    return msg.get_payload(None, True)

            # -----------Function to search for a key value pair
            def search(key, value, con):
                result, data = con.search(None, key, '"{}"'.format(value))
                return data

            # ---------------Function to get the list of emails under this label
            def get_emails(result_bytes):
                msgs = []  # all the email data are pushed inside an array
                for num in result_bytes[0].split():
                    typ, data = con.fetch(num, '(RFC822)')
                    msgs.append(data)
                return msgs

            con = imaplib.IMAP4_SSL(host)
            con.login(username, password)
            con.select('Inbox')

            # --------------fetching emails from a user
            msgs = get_emails(search('FROM', 'noreply@salesforce.com', con))
            Code=""
            for msg in msgs[::-1]:
                for sent in msg:
                    if Code!="":
                        break
                    else:
                        if type(sent) is tuple:
                            content = str(sent[1], 'utf-8')
                            data = str(content)
                            try:
                                indexstart = data.find("ltr")
                                data2 = data[indexstart + 5: len(data)]
                                indexend = data2.find("</div>")
                                indx = data2.find('Verification Code:')
                                Code = data2[indx + 19] + data2[indx + 20] + data2[indx + 21] + data2[indx + 22] + data2[
                                    indx + 23] + data2[indx + 24] + data2[indx + 25] + data2[indx + 26]
                                print(Code)
                                break
                            except UnicodeEncodeError as e:
                                pass

            #---------------To Delete the email from inbox-----------------
            # messages = msgs[0].split(b' ')
            # print("Deleting mails")
            # count = 1
            # for mail in messages:
            #     # mark the mail as deleted
            #     con.store(mail, "+FLAGS", "\\Deleted")
            #     print(count, "mail(s) deleted")
            #     count += 1
            # print("All selected mails has been deleted")
            #
            # # delete all the selected messages
            # con.expunge()
            # # close the mailbox
            # con.close()
            #
            # # logout from the server
            # con.logout()


            #------------Waiting for Verification code email-----------------
            if Code=="":
                time.sleep(7)
                print("Waiting for Verification Code ")

            # -----------------To Capture No Verification Code sent error from Sales Force-------------------------
            try:
                bool1 = driver.find_element_by_id("save").is_displayed()
            except Exception:
                try:
                    bool1 = driver.find_element_by_xpath(
                        "//div/h2[@class='mb12']").is_displayed()
                    if bool1 == True:
                        ErrorFound = driver.find_element_by_xpath(
                            "//div/h2[@class='mb12']").text
                        print(ErrorFound)
                        ErrorFound2 = driver.find_element_by_xpath(
                            "//div[@id='content']/form/p").text
                        print(ErrorFound2)
                        TestResult.append(" Verification code is not able to send from Sales Force due to below\n" + ErrorFound+ErrorFound2)
                        TestResultStatus.append("Fail")
                        driver.close()
                except Exception:
                    pass

            # -----------------Login in Sales Force-------------------------
            time.sleep(2)
            driver.find_element_by_id('emc').send_keys(Code)
            try:
                WebDriverWait(driver, SHORT_TIMEOUT
                              ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

                WebDriverWait(driver, LONG_TIMEOUT
                              ).until_not(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
            except TimeoutException:
                pass

            #-------------------Clicking on Opportunity Tab in Top Menu------------------------
            try:
                driver.find_element_by_xpath("//a[@title='Opportunities']/parent::*").click()
                try:
                    WebDriverWait(driver, SHORT_TIMEOUT
                                  ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

                    WebDriverWait(driver, LONG_TIMEOUT
                                  ).until_not(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
                except TimeoutException:
                    pass
            except Exception:
                PageLoadError = driver.find_element_by_xpath("//span[@jsselect='heading']").text
                print(PageLoadError)
                TestResult.append(
                    " Sales Force site is not able to load. Below error found\n" + PageLoadError)
                TestResultStatus.append("Fail")
                driver.quit()

            #-------------------Clikcing on New--------------------------
            driver.find_element_by_xpath("//a[@title='New']").click()
            try:
                WebDriverWait(driver, SHORT_TIMEOUT
                              ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

                WebDriverWait(driver, LONG_TIMEOUT
                              ).until_not(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
            except TimeoutException:
                pass

            # -------------------Clikcing on Next--------------------------
            driver.find_element_by_xpath("//span[text()='Next']").click()
            try:
                WebDriverWait(driver, SHORT_TIMEOUT
                              ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

                WebDriverWait(driver, LONG_TIMEOUT
                              ).until_not(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
            except TimeoutException:
                pass

            #------------Filling Opportunity details---------------------------

            #--------Opportunity Name-----------
            FieldName="Opportunity Name"
            print(FieldName)
            print(FieldDataSF.get(FieldName))
            driver.find_element_by_xpath("//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[3]/div[1]/div/div/div/input").send_keys(FieldDataSF.get(FieldName))
            time.sleep(2)
            ProjectName=FieldDataSF.get(FieldName)

            # --------Opportunity Close Date-----------
            try:
                FieldName = "Close Date"
                print(FieldName)
                print(FieldDataSF.get(FieldName))

                Duration = int(FieldDataSF.get(FieldName))
                today = datetime.now()
                NewDate = today +timedelta(days=Duration)
                NewDate = NewDate.strftime('%m/%d/%Y')
                if NewDate[0] == "0":
                    Item = ''.join([NewDate[i] for i in range(len(NewDate)) if i != 0])
            except Exception as ed:
                print("excppp")
                print(ed)
                Item="1/20/2022"
                pass
            print(Item)
            driver.find_element_by_xpath(
                "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[4]/div[2]/div/div/div/div/input").send_keys(Item)
            time.sleep(2)

            # --------Opportunity Type-----------
            FieldName = "Opportunity Type"
            for scrolldown in range(1, 10):
                time.sleep(2)
                try:
                    driver.find_element_by_xpath(
                        "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[8]/div[1]/div/div/div/div/div/div/div").click()
                    break
                except Exception:
                    ActionChains(driver).key_down(Keys.PAGE_DOWN).perform()
                    pass
            time.sleep(2)
            for ii3 in range(1, int(FieldDataSF.get(FieldName)) + 1):
                time.sleep(1)
                ActionChains(driver).key_down(Keys.DOWN).perform()
                time.sleep(1)
            ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)
            value1 = driver.find_element_by_xpath(
                "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[8]/div[1]/div/div/div/div/div/div/div/a").text
            print(value1)

            # -------- Process Type-----------
            FieldName = "Process Type"
            for scrolldown in range(1, 10):
                time.sleep(2)
                try:
                    driver.find_element_by_xpath(
                        "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[9]/div[1]/div/div/div/div/div/div/div").click()
                    break
                except Exception:
                    ActionChains(driver).key_down(Keys.PAGE_DOWN).perform()
                    pass
            time.sleep(2)
            for ii3 in range(1, int(FieldDataSF.get(FieldName)) + 1):
                time.sleep(1)
                ActionChains(driver).key_down(Keys.DOWN).perform()
                time.sleep(1)
            ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)
            value1 = driver.find_element_by_xpath(
                "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[9]/div[1]/div/div/div/div/div/div/div/a").text
            print(value1)

            # --------Opportunity Stage-----------
            FieldName = "Stage"
            for scrolldown in range(1, 10):
                time.sleep(2)
                try:
                    driver.find_element_by_xpath(
                        "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[9]/div[2]/div/div/div/div/div/div/div").click()
                    break
                except Exception:
                    ActionChains(driver).key_down(Keys.PAGE_DOWN).perform()
                    pass
            time.sleep(2)
            for ii3 in range(1, int(FieldDataSF.get(FieldName)) + 1):
                time.sleep(1)
                ActionChains(driver).key_down(Keys.DOWN).perform()
                time.sleep(1)
            ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)
            value1 = driver.find_element_by_xpath(
                "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[9]/div[2]/div/div/div/div/div/div/div/a").text
            print(value1)

            # --------Opportunity Sub Stage-----------
            FieldName = "Sub Stage"
            for scrolldown in range(1, 10):
                time.sleep(2)
                try:
                    driver.find_element_by_xpath(
                        "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[10]/div[2]/div/div/div/div/div/div/div").click()
                    break
                except Exception:
                    ActionChains(driver).key_down(Keys.PAGE_DOWN).perform()
                    pass
            time.sleep(2)
            for ii3 in range(1, int(FieldDataSF.get(FieldName)) + 1):
                time.sleep(1)
                ActionChains(driver).key_down(Keys.DOWN).perform()
                time.sleep(1)
            ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)
            value1 = driver.find_element_by_xpath(
                "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[10]/div[2]/div/div/div/div/div/div/div/a").text
            print(value1)

            # --------Opportunity Lead and Referral Source-----------
            FieldName = "Lead and Referral Source"
            for scrolldown in range(1, 10):
                time.sleep(2)
                try:
                    driver.find_element_by_xpath(
                        "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[12]/div[2]/div/div/div/div/div/div/div").click()
                    break
                except Exception:
                    ActionChains(driver).key_down(Keys.PAGE_DOWN).perform()
                    pass
            time.sleep(2)
            for ii3 in range(1, int(FieldDataSF.get(FieldName)) + 1):
                time.sleep(1)
                ActionChains(driver).key_down(Keys.DOWN).perform()
                time.sleep(1)
            ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)
            value1 = driver.find_element_by_xpath(
                "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[12]/div[2]/div/div/div/div/div/div/div/a").text
            print(value1)

            # --------Account Name-----------
            FieldName = "Account Name"
            for scrolldown in range(1, 10):
                time.sleep(2)
                try:
                    driver.find_element_by_xpath(
                        "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[12]/div[1]/div/div/div/div/div/div[1]/div").click()
                    break
                except Exception:
                    ActionChains(driver).key_down(Keys.PAGE_DOWN).perform()
                    pass
            time.sleep(2)
            for ii3 in range(1, int(FieldDataSF.get(FieldName)) + 1):
                time.sleep(1)
                ActionChains(driver).key_down(Keys.DOWN).perform()
                time.sleep(1)
            ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)
            value1 = driver.find_element_by_xpath(
                "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[12]/div[1]/div/div/div/div/div/div[2]/div/ul/li/a/span[2]").text
            print(value1)

            # --------Financial Account-----------
            FieldName = "Financial Account"
            for scrolldown in range (1,10):
                time.sleep(2)
                try:
                    driver.find_element_by_xpath(
                        "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[13]/div[1]/div/div/div/div/div/div[1]/div").click()
                    break
                except Exception:
                    ActionChains(driver).key_down(Keys.PAGE_DOWN).perform()
                    pass
            time.sleep(2)
            for ii3 in range(1,int(FieldDataSF.get(FieldName))+1):
                time.sleep(1)
                ActionChains(driver).key_down(Keys.DOWN).perform()
                time.sleep(1)
            ActionChains(driver).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)
            value1=driver.find_element_by_xpath(
                "//div[@class='modal-body scrollable slds-modal__content slds-p-around--medium']/div/div/div[1]/div/article/div[3]/div/div[1]/div/div/div[13]/div[1]/div/div/div/div/div/div[2]/div/ul/li[1]/a/span[2]").text
            print(value1)

            # --------Opportunity Description-----------
            FieldName = "Description"
            print(FieldName)
            print(FieldDataSF.get(FieldName))
            driver.find_element_by_xpath(
                "//textarea[1]").send_keys(FieldDataSF.get(FieldName))
            time.sleep(2)


            # ------------Submitting Opportunity details---------------------------
            driver.find_element_by_xpath("//div[@class='button-container-inner slds-float_right']/button[3]/span").click()
            time.sleep(10)

            #----------------------Now Navigating to BIDS Application----------------------------
            #-------------------For Login in BIDS-------------------
            driver.get("https://beneficienttest.appiancloud.com/suite/")
            driver.find_element_by_id("un").send_keys("neeraj.kumar")
            driver.find_element_by_id("pw").send_keys("Crochet@786")
            driver.find_element_by_xpath("//input[@type='submit']").click()

            #---------------------------Verify Transactions page-----------------------------
            PageName="Transactions"
            Ptitle1="Transaction Listing "
            driver.find_element_by_xpath("//*[@title='"+PageName+"']").click()
            start = time.time()
            try:
                WebDriverWait(driver, SHORT_TIMEOUT
                              ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

                WebDriverWait(driver, LONG_TIMEOUT
                              ).until_not(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
            except TimeoutException:
                pass
            try:
                time.sleep(2)
                bool1 = driver.find_element_by_xpath(
                    "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[1]").is_displayed()
                if bool1 == True:
                    ErrorFound1 = driver.find_element_by_xpath(
                        "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[1]").text
                    print(ErrorFound1)
                    driver.find_element_by_xpath(
                        "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[2]/div/button").click()
                    TestResult.append(PageName + " not able to open\n" + ErrorFound1)
                    TestResultStatus.append("Fail")
                    bool1 = False
            except Exception:
                try:
                    time.sleep(2)
                    bool2 = driver.find_element_by_xpath(
                        "//div[@class='MessageLayout---message MessageLayout---error']").is_displayed()
                    if bool2 == True:
                        ErrorFound2 = driver.find_element_by_xpath(
                            "//div[@class='MessageLayout---message MessageLayout---error']/div/p").text
                        print(ErrorFound2)
                        TestResult.append(PageName + " not able to open\n" + ErrorFound2)
                        TestResultStatus.append("Fail")
                        bool2 = False
                except Exception:
                    pass
                pass
            time.sleep(1)
            try:
                PageTitle1 = driver.find_element_by_xpath("//div[@class='ContentLayout---content_layout']/div[2]/div/div/div/div/div/div[1]/div/div/div").text
                print(PageTitle1)
                assert Ptitle1 in PageTitle1, PageName + " not able to open"
                TestResult.append(PageName + " page Opened successfully")
                TestResultStatus.append("Pass")
            except Exception:
                TestResult.append(PageName +     " page not able to open")
                TestResultStatus.append("Fail")
            print()
            stop = time.time()
            TimeString = stop - start
            #---------------------------------------------------------------------------------

            try:
                print()
                TotalItem=driver.find_element_by_xpath(
                    "//div[@class='ContentLayout---content_layout']/div[2]/div/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/span[2]").text
                print("TotalItem "+ TotalItem)
                substr = "of"
                x = TotalItem.split(substr)
                string_name = x[0]
                TotalItemAfterOf= x[1]
                abc = ""
                countspace = 0
                for element in range(0, len(string_name)):
                    if string_name[(len(string_name) - 1) - element] == " ":
                        countspace = countspace + 1
                        if countspace == 2:
                            break
                    else:
                        abc = abc + string_name[(len(string_name) - 1) - element]
                abc = abc[::-1]
                TotalItemBeforeOf=abc
                print("TotalItemAfterOf " + TotalItemAfterOf)
                print("TotalItemBeforeOf " + TotalItemBeforeOf)

                #----------------Searching the Project from Sales Force--------------------
                ProejctTOClick = ProjectName
                for waiting in range(1,3):
                    print("Waiting Iteration "+str(waiting))
                    time.sleep(60)

                    IterateNo = int(TotalItemAfterOf) / int(TotalItemBeforeOf)
                    if IterateNo.is_integer()==True:
                        #print("Yes Integer")
                        IterateNo=IterateNo-1
                        pass
                    else:
                        #print("No Integer")
                        print(str(float(IterateNo)))
                        IterateNo = math.ceil(float(IterateNo))
                        print(IterateNo)
                        print()

                    loopbreak=0
                    for ii5 in range(1, IterateNo+1):
                        if loopbreak==0:
                            if ii5 >1:
                                try:
                                    driver.find_element_by_xpath(
                                        "//div[@class='ContentLayout---content_layout']/div[2]/div/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/span[4]/a[1]").click()
                                    try:
                                        WebDriverWait(driver, SHORT_TIMEOUT
                                                      ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

                                        WebDriverWait(driver, LONG_TIMEOUT
                                                      ).until_not(
                                            EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
                                    except TimeoutException:
                                        pass
                                except Exception as q1:
                                    print(q1)
                                    pass

                            RowsInv = driver.find_elements_by_xpath(
                                "//div[@class='ContentLayout---content_layout']/div[2]/div/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody/tr")
                            for ii3 in range(1, len(RowsInv)+1):
                                ProjectNameText = driver.find_element_by_xpath(
                                    "//div[@class='ContentLayout---content_layout']/div[2]/div/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody/tr[" + str(
                                        ii3) + "]/td[1]/div/p/a").text
                                print(ProjectNameText)
                                if ProjectNameText==ProejctTOClick:
                                    loopbreak=1
                                    PageName=ProejctTOClick
                                    driver.find_element_by_xpath("//div[@class='ContentLayout---content_layout']/div[2]/div/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div[1]/div[2]/table/tbody/tr/td[1]/div/p/a[text()='"+ProjectNameText+"']").click()
                                    try:
                                        WebDriverWait(driver, SHORT_TIMEOUT
                                                      ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

                                        WebDriverWait(driver, LONG_TIMEOUT
                                                      ).until_not(
                                            EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
                                    except TimeoutException:
                                        pass
                                    try:
                                        time.sleep(2)
                                        bool1 = driver.find_element_by_xpath(
                                            "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[1]").is_displayed()
                                        if bool1 == True:
                                            ErrorFound1 = driver.find_element_by_xpath(
                                                "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[1]").text
                                            print(ErrorFound1)
                                            driver.find_element_by_xpath(
                                                "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[2]/div/button").click()
                                            TestResult.append(PageName + " not able to open\n" + ErrorFound1)
                                            TestResultStatus.append("Fail")
                                            bool1 = False
                                    except Exception:
                                        try:
                                            time.sleep(2)
                                            bool2 = driver.find_element_by_xpath(
                                                "//div[@class='MessageLayout---message MessageLayout---error']").is_displayed()
                                            if bool2 == True:
                                                ErrorFound2 = driver.find_element_by_xpath(
                                                    "//div[@class='MessageLayout---message MessageLayout---error']/div/p").text
                                                print(ErrorFound2)
                                                TestResult.append(PageName + " not able to open\n" + ErrorFound2)
                                                TestResultStatus.append("Fail")
                                                bool2 = False
                                        except Exception:
                                            pass
                                        pass
                                    break
                        else:
                            break

                # ------------------------clicking Transaction ID--------------------------------
                PageName = "Transaction ID"
                driver.find_element_by_xpath("//div[@class='ContentLayout---content_layout']/div[4]/div[2]/div/div/div[2]/div/div/table/tbody/tr/td[2]/div/p/a").click()
                try:
                    WebDriverWait(driver, SHORT_TIMEOUT
                                  ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

                    WebDriverWait(driver, LONG_TIMEOUT
                                  ).until_not(
                        EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
                except TimeoutException:
                    pass
                try:
                    time.sleep(2)
                    bool1 = driver.find_element_by_xpath(
                        "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[1]").is_displayed()
                    if bool1 == True:
                        ErrorFound1 = driver.find_element_by_xpath(
                            "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[1]").text
                        print(ErrorFound1)
                        driver.find_element_by_xpath(
                            "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[2]/div/button").click()
                        TestResult.append(PageName + " not able to open\n" + ErrorFound1)
                        TestResultStatus.append("Fail")
                        bool1 = False
                except Exception:
                    try:
                        time.sleep(2)
                        bool2 = driver.find_element_by_xpath(
                            "//div[@class='MessageLayout---message MessageLayout---error']").is_displayed()
                        if bool2 == True:
                            ErrorFound2 = driver.find_element_by_xpath(
                                "//div[@class='MessageLayout---message MessageLayout---error']/div/p").text
                            print(ErrorFound2)
                            TestResult.append(PageName + " not able to open\n" + ErrorFound2)
                            TestResultStatus.append("Fail")
                            bool2 = False
                    except Exception:
                        pass
                    pass
                time.sleep(1)

                #-------------clicking Edit Key Transaction Details--------------------------------
                driver.find_element_by_xpath("//p/a[text()='Edit Key Transaction Details']").click()
                PageName="Edit Key Transaction Details"
                try:
                    WebDriverWait(driver, SHORT_TIMEOUT
                                  ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

                    WebDriverWait(driver, LONG_TIMEOUT
                                  ).until_not(
                        EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
                except TimeoutException:
                    pass
                try:
                    time.sleep(1)
                    bool1 = driver.find_element_by_xpath(
                        "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[1]").is_displayed()
                    if bool1 == True:
                        ErrorFound1 = driver.find_element_by_xpath(
                            "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[1]").text
                        print(ErrorFound1)
                        driver.find_element_by_xpath(
                            "//div[@class='appian-context-ux-responsive']/div[4]/div/div/div[2]/div/button").click()
                        TestResult.append(PageName + " not able to open\n" + ErrorFound1)
                        TestResultStatus.append("Fail")
                        bool1 = False
                except Exception:
                    try:
                        time.sleep(1)
                        bool2 = driver.find_element_by_xpath(
                            "//div[@class='MessageLayout---message MessageLayout---error']").is_displayed()
                        if bool2 == True:
                            ErrorFound2 = driver.find_element_by_xpath(
                                "//div[@class='MessageLayout---message MessageLayout---error']/div/p").text
                            print(ErrorFound2)
                            TestResult.append(PageName + " not able to open\n" + ErrorFound2)
                            TestResultStatus.append("Fail")
                            bool2 = False
                    except Exception:
                        pass
                    pass
                time.sleep(1)
                # -------------Fetching Key Transaction Details--------------------------------


            except Exception:
                pass

        except Exception as Mainerror:
            stringMainerror=repr(Mainerror)
            if stringMainerror in "InvalidSessionIdException('invalid session id', None, None)":
                pass
            else:
                TestResult.append(stringMainerror)
                TestResultStatus.append("Fail")

    else:
        print()
        print("Test Case skipped as per the Execution sheet")
        skip = "Yes"

        # -----------To add Skipped test case details in PDF details sheet-------------
        ExcelFileName = "FileName"
        loc = (path+'PDFFileNameData/' + ExcelFileName + '.xlsx')
        wb = openpyxl.load_workbook(loc)
        sheet = wb.active
        check = TestName

        for i in range(1, 100):
            if sheet.cell(i, 1).value == check:
                sheet.cell(row=i, column=5).value = "Skipped"
                wb.save(loc)
        # ----------------------------------------------------------------------------


