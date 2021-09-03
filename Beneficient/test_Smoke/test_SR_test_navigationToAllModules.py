import os
import smtplib
from email.message import EmailMessage

import openpyxl
import pytest


@pytest.mark.smoke
def test_ReportSend_AllModulesVerify():
    ExcelFileName = "FileName"
    loc = ('C:/BIDS/beneficienttest/Beneficient/PDFFileNameData/' + ExcelFileName + '.xlsx')
    wb=openpyxl.load_workbook(loc)
    sheet = wb.active
    PDFName=sheet.cell(1, 1).value
    print(PDFName)

    msg=EmailMessage()
    # mention failure and pass status ---------------------------

    msg['Subject']='Automation Test Report'
    msg['From']='Neeraj'
    msg['To']='neeraj.kumar@crochetech.com'
    #msg.set_content("Test email from Neeraj")

    with open('C:/BIDS/beneficienttest/Beneficient/test_Smoke/EmTemp.txt') as myfile:
        data=myfile.read()

    with open('C:/BIDS/beneficienttest/Beneficient/test_Smoke/'+PDFName, 'rb') as f:
        file_data = f.read()
        file_name = f.name

    msg.set_content(data)
    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    server=smtplib.SMTP_SSL('smtp.gmail.com',465)
    server.login("neeraj.kumar@bitsinglass.com","Motorola@408")

    server.send_message(msg)
    print("Test Report sent")
    os.remove('C:/BIDS/beneficienttest/Beneficient/test_Smoke/'+PDFName)
    server.quit()
