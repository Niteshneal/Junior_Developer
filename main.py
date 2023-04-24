import os
from selenium import webdriver
import re
import csv
import time
import xlsxwriter

def historical():
    driver = webdriver.Chrome(executable_path='C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe')

    url = driver.get('https://finance.yahoo.com/quote/BTC-EUR/history?p=BTC-EUR')
    driver.maximize_window()
    time.sleep(7)

    drop_down = driver.find_element_by_xpath("//div[@class='Pos(r) D(ib) C($linpaintkColor) Cur(p)']").click()

    start_Date = driver.find_element_by_xpath("//div[@class='Mb(10px)']/*[name()='input']").send_keys('30/08/2021')

    #end_Date = driver.find_element_by_xpath("//span[@class='C($tertiaryColor) Fz(12px) Mend(15px)']/*[name()='input']").send_keys('08/09/2021')

    finish = driver.find_element_by_xpath("//button[@class=' Bgc($linkColor) Bdrs(3px) Px(20px) Miw(100px) Whs(nw) Fz(s) Fw(500) C(white) Bgc($linkActiveColor):h Bd(0) D(ib) Cur(p) Td(n)  Py(9px) Miw(80px)! Fl(start)']").click()

    process = driver.find_element_by_xpath("//button[@data-reactid='25']").click()

    time.sleep(2)

    upLoad = driver.find_element_by_xpath("//span[@class='Fl(end) Pos(r) T(-6px)']").click()
    time.sleep(10)

    driver.close()


def creation():
    Date, BTC_Closing_Value = [], []
    row = 2

    locate = open(r'C:\Users\stuti sharma\Downloads\BTC-EUR.csv', 'r')
    csv_btc = csv.reader(locate)

    for i in csv_btc:
        Date.append(i[0])
        BTC_Closing_Value.append(i[4])

    print(Date)
    print(BTC_Closing_Value)

#This is the csv version

    rename_file = open('eur_btc_rates.csv', 'w')
    csv_btc_write = csv.writer(rename_file)
    csv_btc_write.writerow(Date)
    csv_btc_write.writerow(BTC_Closing_Value)

#making it using xlsx file, to have a better view in the excel sheet

    project = xlsxwriter.Workbook("eur_btc_rates.xlsx")

    project_sheet = project.add_worksheet("BTC-EUR-Historical-Data")
    project_sheet.write('A1', 'Date')
    project_sheet.write('B1', 'BTC_Closing_Value')

    for i in range(len(Date)):
        project_sheet.write('A' +str(row), Date[i])
        project_sheet.write('B' + str(row), BTC_Closing_Value[i])
        row+=1

    project.close()

historical()
creation()
