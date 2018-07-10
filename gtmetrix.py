from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import openpyxl
#import urllib
#import requests, bs4

wb = openpyxl.load_workbook('C:/Git/Montaigne/format.xlsx')  # For Opening workbook
sheet = wb.get_sheet_by_name('Dallas')
u1 = sheet['A3'].value              # To get url link from Excel sheet
sheet['B3'] = 'B(78%)'
sheet['B4'] = 'B(54%)'
sheet['B5'] = 'B(88%)'
sheet['B6'] = 'B(52%)'
sheet['B7'] = 'B(99%)'

wb.save('C:/Git/Montaigne/format.xlsx')
email = "shankarnag151@gmail.com"   # For Login into the GTmetrix website
passwd = "8861415188"

dr = webdriver.Chrome()
dr.maximize_window()
dr.get('https://gtmetrix.com/')

linkElem = dr.find_element_by_link_text('Log In')
linkElem.click()
elem = dr.find_element_by_name("email")     # Finding Email field in GTmetrix website.
elem.send_keys(email)
elem = dr.find_element_by_name("password")  # Finding Password field in GTmetrix website.
elem.send_keys(passwd)
elem.submit()


elem = dr.find_element_by_name("url")
elem.send_keys(u1)
linkElem = dr.find_element_by_link_text('Analyze')
linkElem.click()

linkElem = dr.find_element_by_link_text('Log Out')
linkElem.click()                           
dr.close()

#elems[0].getText()
#r1 = elems.getText('report-score-percent')
#print (r1)

#if __name__ == "__main__":
#    unittest.main()
