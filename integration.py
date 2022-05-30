#!/usr/bin/env python
#
# Very basic example of using Python 3 and IMAP to iterate over emails in a
# gmail folder/label.  This code is released into the public domain.
#
# This script is example code from this blog post:
# http://www.voidynullness.net/blog/2013/07/25/gmail-email-with-python-via-imap/
#
# This is an updated version of the original -- modified to work with Python 3.4.
#
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.webdriver import FirefoxProfile
from datetime import date, timedelta, datetime
import sys
import imaplib
import getpass
import email
import email.header
import os
import csv
import glob
import time
import requests
import json
import requests
from openpyxl import load_workbook, Workbook
import pandas as pd
import numpy
import shutil
import smtplib


from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from random import randint

Initial_path = 'E:\\pos\\dirname\\attachment'

# filename = max([os.path.join(Initial_path, f) for f in os.listdir(Initial_path)], key=os.path.getctime)
# print(filename)
# sys.exit()

chrome_options = Options()
chrome_options.add_argument('--dns-prefetch-disable')
chrome_options.add_argument("no-sandbox") 
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("download.default_directory=E:\\pos\\dirname\\attachment")
#prefs = {'download.default_directory' : 'D:/wamp64/www/projects/dirname/attachments'}
prefs = {'download.default_directory' : 'E:\\pos\\dirname\\attachment'}
#prefs = {}
chrome_options.add_experimental_option('prefs', prefs)

#driver = webdriver.Chrome(chrome_options=chrome_options)
#driver = webdriver.Firefox()
driver = webdriver.Chrome(chrome_options=chrome_options)

#try:
	
try:
	driver.get('https://app.posible.in/')
	time.sleep(2)
	a = driver.find_elements_by_xpath('.//input')
	#print(a[0].text)
	a[0].send_keys('login_id')
	a[1].send_keys('password')

	#driver.find_element_by_id('Password').send_keys('xxxxx')

	#a[2].click()
	time.sleep(1)
	#c2 =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button-primary')))
	#c2 = trlist = driver.find_elements_by_xpath(".//button[@class = 'button-primary']")
	#driver.execute_script("arguments[0].click()", c2)
	driver.execute_script("document.getElementsByClassName('button-primary')[0].click()")
except Exception as err:
	print(str(err))
	driver.quit()
	

print ("Login successfull")


# rt = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'root')))
# time.sleep(5)
# trlist = driver.find_elements_by_xpath(".//tr[@class = 'reportTabletr']")
# tdlist = trlist[1].find_elements_by_xpath("td")
# print(tdlist[len(tdlist)-1].text)
# print(tdlist[len(tdlist)-1].click())

#sys.exit()


time.sleep(8)
strslists = driver.find_elements_by_xpath(".//div[@class = 'list-row']")
# print(len(strslists))
driver.execute_script("arguments[0].click()", strslists[0])
# sys.exit()
time.sleep(10)
#driver.get('https://URL/reports')
driver.execute_script("document.getElementsByClassName('btn-menubar')[0].click()")
time.sleep(1)
mitems = driver.find_elements_by_xpath(".//div[@class = 'menu-item-text']")
for t in range(len(mitems)):
	litext = mitems[t].text.strip()
	if(litext == 'Reports'):
		driver.execute_script("arguments[0].click()", mitems[t])
	print(litext)

time.sleep(2)	
mch = driver.find_element_by_xpath(".//div[@class = 'menu-item-child-container']")
mitems = mch.find_elements_by_xpath("div")
print(1111)
print(len(mitems))
for t in range(len(mitems)):
	litext = mitems[t].text.strip()
	if(litext == 'Report Options'):
		print(litext)
		driver.execute_script("arguments[0].click()", mitems[t])
	print(litext)

time.sleep(1)	
invrprt = driver.find_element_by_link_text("Invoice Detail Report")
driver.execute_script("arguments[0].click()", invrprt)

time.sleep(4)
#driver.execute_script("document.getElementsByName('input')[1].setAttribute('value','2021-09-23')")
#driver.find_element_by_xpath("//input[@type='date']").send_key('2021-09-23')
# dtimps = driver.find_elements_by_xpath(".//input")
# print(222)
# driver.execute_script("arguments[0].setAttribute('value',arguments[1])",dtimps[1], '2021-09-23')
# print(333)

# boxbody = driver.find_element_by_xpath(".//div[@class = 'box-body']")
boxbody = Select(driver.find_elements_by_css_selector('select'))
# boxbody[0].select_by_visible_text('All Stores').click()
print('in select search')
print(len(boxbody))
sys.exit()

driver.execute_script("document.getElementsByClassName('fa-download')[0].click()")



time.sleep(2)

#c4 =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'csvdownload')))
#driver.execute_script("arguments[0].click()", c4)

print("downloading started")

time.sleep(10)

#sys.exit()






#driver.close()

#except Exception as err:
#	print('execption occured!')
	#driver.close()
	#driver.quit()


time.sleep(10)

#sys.exit()
filename = max([os.path.join(Initial_path, f) for f in os.listdir(Initial_path)], key=os.path.getctime)
print(filename)

csvFile = filename

fileSalewise = 	filename	 
print(fileSalewise)
df = pd.read_csv(fileSalewise,dtype={"customermobile": str})

sheetinjson = df.to_numpy()

orders = []
rows = []
itemJson = {}

#clean_dict = filter(lambda k: not isnan(sheetinjson[k]), sheetinjson)
somewithnextrow = 0
orderitems = []
allitems = {}
i = 0
for js in sheetinjson:
	print(js[0])	

#print(refnoindex)
#billno = str(js[3])
#print(str(type(billno))+': '+str(billno))
#if billno == "nan":
	#print('this is nan : '+str(billno))
#	print('')

	#i = i+1

driver.quit()

emailbody = MIMEMultipart('alternative')
yDate = datetime.strftime(datetime.now() - timedelta(1), '%Y-%m-%d')
emailbody['Subject'] = "Your subject to send email"
emailbody['From'] = "<From email id>"
emailbody['To'] = "<To email id>"


msg = "We synced last one day  orders report <br>"
part2 = MIMEText(msg, 'html')
emailbody.attach(part2)
server = smtplib.SMTP('smtp.gmail.com', 587)

server.starttls()

toaddr=["TO_EMAIL1","TO_EMAIL2"]
	
server.login("smtp_LOGIN_ID", "smtp_LOGIN_pwd")

server.sendmail("<Default to email addr>", toaddr , emailbody.as_string())
print("email line executed")
server.quit()

### move file
# dest_folder = r'E:\pos\soultree\storeage'
# m_fir = r'E:\pos\dirname\storeage\attachment'
# nfilename = m_fir+'\\'+str(datetime.datetime.now().time() )+ '.csv'
# os.rename(filename,nfilename) 
# files = [nfilename]
# for f in files:
#     shutil.copy(f, dest_folder)

#os. remove(filename)

# except Exception as err:
# 	print('execption occured!')
# 	driver.close()
# 	driver.quit()
					
					
