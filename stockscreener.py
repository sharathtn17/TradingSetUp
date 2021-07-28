import os
import subprocess
import smtplib
import sys
import xlsxwriter
import string
#from xlrd import Workbook
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
URL="abc"
print("start")
input=sys.argv[1].strip()
file=open("properties.txt","r")
for line in file:
    method=line.split(":",1)
    print(input+"PARM")
    print(method[0].strip()+"URL")
    m=method[0].strip()
    if ( input == m):
       print(method[1]+" URL")
       URL=method[1]
       print(method[1]+"URL TUR####")
    else:
        print("NOT FOUND")

option = webdriver. ChromeOptions()
option. add_argument('headless')
browser=webdriver.Chrome("chromedriver.exe",options=option)
#browser=webdriver.Chrome()
browser.get(URL)


print(URL)
print(input)
# Give the location of the file
# Workbook() takes one, non-optional, argument 
# which is the filename that we want to create.
workbook = xlsxwriter.Workbook(input+'.xlsx')
sheet1 = workbook.add_worksheet(input)
# add_sheet is used to create sheet.

#sheet1 = wb.add_sheet(method[0])
sheet1.write(0, 0, "STOCKS")
sheet1.write(0, 1, "Price")
sheet1.write(0, 2, "ChangeinPercentage")
sheet1.write(0, 3, "Volume")
#browser.find_element(by=By.XPATH,value="//input[@id='sso_username']").send_keys('sharath')
browser.implicitly_wait(2000)

c=1
print("findinf ele")
browser.implicitly_wait(0.5)
disab=browser.find_elements_by_class_name("paginate_button next disabled")
print(len(disab))
pagination=browser.find_elements_by_xpath("//li[@class='paginate_button ']")
pagecount=len(pagination)
if len(disab) == 0:
  pagecount+=1
  print(pagecount)

page=1

while True:
    element=1
    count=browser.find_elements_by_xpath("//table[@class='table table-striped scan_results_table dataTable no-footer']/tbody/tr")
    print(count)
    num=len(count)
    for i in count:
      print("###########3")
      StockName=browser.find_element_by_xpath("//table[@class='table table-striped scan_results_table dataTable no-footer']/tbody/tr["+str(element)+"]/td[2]").text
      Price=browser.find_element_by_xpath("//table[@class='table table-striped scan_results_table dataTable no-footer']/tbody/tr["+str(element)+"]/td[6]").text
      Volume=browser.find_element_by_xpath("//table[@class='table table-striped scan_results_table dataTable no-footer']/tbody/tr["+str(element)+"]/td[7]").text
      ChangeinPercentage=browser.find_element_by_xpath("//table[@class='table table-striped scan_results_table dataTable no-footer']/tbody/tr["+str(element)+"]/td[5]").text
      sheet1.write(c, 0, StockName)
      sheet1.write(c, 1, Price)
      sheet1.write(c, 2, ChangeinPercentage)
      sheet1.write(c, 3, Volume)
      c+=1
      element+=1
      print(StockName)
      print(c)
     
    if(page >=pagecount):
       break  
    else:
       print("page###")
       print(page)
       print("page@@@")
       ele=browser.find_element_by_xpath("//li[@class='paginate_button ']["+str(page)+"]/a")
       print(ele.text)
       ele.click()
       browser.implicitly_wait(10000)
    
    
    page+=1


workbook.close()
#print(num)
#os.system('sendm.py input')
#subprocess.Popen("sendm.py input", shell=True)