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

Initial_path = 'E:\\pos\\soultree\\attachment'

# filename = max([os.path.join(Initial_path, f) for f in os.listdir(Initial_path)], key=os.path.getctime)
# print(filename)
# sys.exit()

chrome_options = Options()
#chrome_options.add_argument("user-data-dir=D:\\E-Drive\\projects\\python\\doubledown\\cache" )
chrome_options.add_argument('--dns-prefetch-disable')
chrome_options.add_argument("no-sandbox") 
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("download.default_directory=E:\\pos\\soultree\\attachment")
#prefs = {'download.default_directory' : 'D:/wamp64/www/projects/taghash/python/posintegration/whataburger/attachments'}
prefs = {'download.default_directory' : 'E:\\pos\\soultree\\attachment'}
#prefs = {}
chrome_options.add_experimental_option('prefs', prefs)

#driver = webdriver.Chrome(chrome_options=chrome_options)
#driver = webdriver.Firefox()
driver = webdriver.Chrome(chrome_options=chrome_options)

try:
	
	try:
		driver.get('https://app.posible.in/')
		time.sleep(2)
		a = driver.find_elements_by_xpath('.//input')
		#print(a[0].text)
		#a[0].send_keys('mloyal.wab@soultree.in')
		a[0].send_keys('Jagan@soultree.in')
		a[1].send_keys('9717431664')

		#driver.find_element_by_id('username').send_keys('ch.vinodkumar1984@gmail.com')



		#driver.find_element_by_id('Password').send_keys('9441253007')

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

	#driver.get('https://soultree.in/dashboards/analytics/reports')
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
		

		outletindex = 0
		dateindex = 8
		invoiceNoindex = 6
		nameindex = 3
		phoneindex = 4
		
		# timeindex = ''
		# cityindex = ''
		# orderstatusindex = ''
		
		# billtypeindex = ''
		# sourceindex = ''
		# pymtmodeindex = ''
		# subtotalindex = 0
		# totalDiscountindex = 0
		# loyaltypointsindex = 0
		# sgstindex = 0
		# cgstindex = 0	
		# netpaybleindex = 0
		
		
		# emailindex = ''
		# addressindex = ''
		# areaindex = ''

				

		# 		j = j+1
		# 	print('column names index has been set')


		#if i>0 :
		invoiceNo = str(js[6])
		# city = str(js[cityindex])
		# orderstatus = str(js[orderstatusindex])
		tt = str(js[dateindex])
		
		#sys.exit()
		#btime = str(js[timeindex]).strip()
		print(invoiceNo)
		#print(tt)
		#sys.exit()
		
		if 1 :
			print(invoiceNo.strip())
			
			#billtime = ws.cell(row = i, column = 2).value
			#billtime = datetime.strptime(billtime, '%H:%M %p')
			#billtime = billtime.strftime('%H:%M:%S')
			#bd = datetime.strptime(tt,'%m/%d/%Y %I:%M:%S %p')
			#billdttime = bd.strftime('%Y-%m-%d %H:%M:%S')
			#buildDate = bd.strftime('%Y-%m-%d')
			buildDate = tt 
			billdttime = tt+' 00:00:00'
			print(billdttime)
			print(buildDate)
			
			outlet = str(js[outletindex])

			store_code = 'ONLINE'
			
			store_code = 'soultree'

			


		
			# billtype = str(js[billtypeindex])
			# pymtmode = str(js[pymtmodeindex])
			# subtotal = str(js[subtotalindex])
			# loyaltypoints = str(js[loyaltypointsindex])
			# cgst = str(js[cgstindex])
			# sgst = str(js[sgstindex])

			bill_discount =0 

			totaltax = 0
			# if (cgst is None) and (sgst is None):
			# 	print('gst none type')
			# else:
			# 	totaltax = round(float(cgst) + float(sgst), 2)
			

			# totalDiscount = str(js[totalDiscountindex])
			# netpayble = str(js[netpaybleindex])
			#netpaid = str(js[netpaidindex])
			
			print('INV'+invoiceNo)

			phone = str(js[phoneindex])
			if phone == 'nan':
				phone=''

			print(phone)
			#sys.exit()

			name = str(js[nameindex])
			if name == 'nan':
				name=''

			# email = str(js[emailindex])
			# if email == 'nan':
			# 	email=''

			# address = str(js[addressindex])
			# if address == 'nan':
			# 	address=''

			# area = str(js[areaindex])
			# if address == 'nan':
			# 	address=''
			
			# billgross = round(float(netpayble) - totaltax, 2)
			# print('bill gross amount : '+str(billgross)+' for oid:'+invoiceNo)
		
		
			orderIdsFile = 'logs/orderIds_'+str(date.today().strftime('%m-%d-%y'))+'.txt'
			isOrderidExist = 0

			if glob.glob(orderIdsFile):
				print("")
			else :	
				open(str(orderIdsFile), 'w+', encoding="utf-8")

			if invoiceNo in open(orderIdsFile).read():								
				print("Order Id "+invoiceNo+" already inserted")
				isOrderidExist = isOrderidExist + 1
			else :
				print("Not matched order id "+invoiceNo)	

			print("Order id exist check - "+str(isOrderidExist))	

			# if pymtmode == '' or pymtmode is None:
			# 	pymtmode = 'cash'

			if phone is None:
				print('')
			else:
				slen = len(phone)
				#phone = phone[-10:slen]
				#phone = phone[-10:slen]
			print('phone : '+str(phone))

			phone = phone.replace(".0","")

			totalamt = 0
			bill_tax = 0
			### pos integration 
			if isOrderidExist == 0 :	
				print("insert proces for order id : "+invoiceNo)
				orderJson = { 
					"objClass": [{									
						"bill_cancel_against": "",
			            "bill_cancel_amount": "",
			            "bill_cancel_date": "",
			            "bill_cancel_reason": "",
			            "bill_cancel_time": "",
			            "bill_date": buildDate	,
			            "bill_discount": "0",
			            "bill_discount_per": "",
			            "bill_grand_total": 0,
			            "bill_gross_amount": 0,
			            "bill_modify": "",
			            "bill_modify_datetime": "",
			            "bill_modify_reason": "",
			            "bill_net_amount": 0,
			            "bill_no": invoiceNo,
			            "bill_remarks1": "",
			            "bill_remarks2": "",
			            "bill_remarks3": "",
			            "bill_remarks4": "",
			            "bill_remarks5": "",
			            "bill_round_off_amount": "",
			            "bill_service_tax": "",
			            "bill_status": "New",
			            "bill_tax": 0,
			            "bill_tender_type": "Cash",
			            "bill_time": billdttime,
			            "bill_transaction_type": "New",
			            "bill_transcation_no": "",
			            "bill_type": "Dinein",
			            "customer_address": "Dummy",
			            "customer_area": "",
			            "customer_city": "",
			            "customer_code": "",
			            "customer_doa": "",
			            "customer_dob": "",
			            "customer_email": "",
			            "customer_fname": name,
			            "customer_gender": "",
			            "customer_lname": "",
			            "customer_mobile": phone,
			            "customer_remarks1": "",
			            "customer_remarks2": "",
			            "customer_remarks3": "",
			            "customer_remarks4": "",
			            "customer_remarks5": "",
			            "customer_state": "",
			            "ext_param1": "",
			            "ext_param2": "",
			            "ext_param3": "",
			            "ext_param4": "",
			            "ext_param5": "",               
						"output": [],
						"store_code": store_code,
			            "voucher_code": "",
			            "voucher_type": "",
			            "voucher_value": ""
					}]
					}

				p = 0 
				for iis in sheetinjson :
					invoiceNovar1 = str(iis[6])

					if (invoiceNo == invoiceNovar1):

						itemcode = str(iis[15])
						barcode = str(iis[16])
						itemname = str(iis[17])
						brand = str(iis[12])
						hsn = str(iis[11])
						qty = str(iis[38])
						itm_tax = round(float(iis[51]),2)
						itm_discount = round(float(iis[52]),2)
						amt = round(float(iis[53]),2)
						mrp = round(float(iis[35]),2)
						serial_no = itemcode+hsn
						serial_no = serial_no.replace(".0","")

						items = {}
						items["item_serial_no"] = itemcode+hsn
						items["item_barcode"] =  ""
						items["item_code"] = itemcode
						items["item_name"] = itemname
						items["item_rate"] = mrp
						items["item_net_amount"] = amt
						items["item_gross_amount"] = amt
						items["item_quantity"] = qty
						items["item_discount_per"] = ""
						items["item_discount"] = itm_discount
						items["item_tax"] = itm_tax
						items["item_service_tax"] = ""
						items["item_brand_code"] = ""
						items["item_brand_name"] = brand
						items["item_category_code"] = ""
						items["item_category_name"] = ""
						items["item_group"] = ""
						items["item_group_name"] = ""
						items["item_color_code"] = ""
						items["item_color_name"] = ""
						items["item_size_code"] = ""
						items["item_size_name"] = ""
						items["item_sub_category_code"] = ""
						items["item_sub_category_name"] = ""
						items["item_status"] = ""
						items["item_department_code"] = ""
						items["item_department_name"] = ""
						items["item_remarks1"] = ""
						items["item_remarks2"] = ""
						items["item_remarks3"] = ""
						items["item_remarks4"] = ""
						items["item_remarks5"] = ""   

						totalamt = totalamt + amt 
						bill_tax = bill_tax + itm_tax                              
						bill_discount = bill_discount + itm_discount
						
						orderJson['objClass'][0]['output'].append(items)

					p = p + 1
					
				
				
				orderJson['objClass'][0]['bill_grand_total'] = totalamt
				orderJson['objClass'][0]['bill_net_amount'] = totalamt
				orderJson['objClass'][0]['bill_gross_amount'] = totalamt
				orderJson['objClass'][0]['bill_tax'] = bill_tax
				orderJson['objClass'][0]['bill_discount'] = bill_discount
				

				orderJsonfile = 'logs/orderJson_'+str(date.today().strftime('%m-%d-%y'))+'.txt'
				if glob.glob(orderJsonfile):
					print("")
				else :	
					open(str(orderJsonfile), 'w+', encoding="utf-8")

				f = open(orderJsonfile,'a')
				f.write(json.dumps(orderJson, indent=4, sort_keys=True))
				f.close()
				#sys.exit()
				headers = {'Content-Type': 'application/json','userid':'mob_usr','pwd':'@8e7d2515-f7b9-4bbe-93ac-59f5fae6c1f7',"Content-Length": str(len(orderJson)),"Accept": "*/*"}       

				r = requests.post("http://Soultreews.mloyalcapture.com/Service.svc/INSERT_BILLING_DATA_ACTION", data=json.dumps(orderJson), headers=headers)    
				print(r.content)

				responseFile = 'logs/response_'+str(date.today().strftime('%m-%d-%y'))+'.txt'
				if glob.glob(responseFile):
					print("")
				else :	
					open(str(responseFile), 'w+', encoding="utf-8")
					
				appendresponse = open(responseFile, "a")
				
				appendresponse.write(str(r.content)+',,,, order id - '+str(invoiceNo))
				appendresponse.write("\n")
				appendresponse.close()

				appendOrder = open(orderIdsFile, "a")
				# write line to output file
				appendOrder.write(invoiceNo)
				appendOrder.write("\n")
				appendOrder.close()

				#sys.exit()

		else :
			#failedFile = 'logs/failed_'+str(date.today().strftime('%m-%d-%y'))+'.txt'
			#if glob.glob(failedFile):
			print("")
			#else :	
			#	open(str(failedFile), 'w+', encoding="utf-8")				
			#appendresponse = open(failedFile, "a")			
			#appendresponse.write(',,,, Failed order id - '+str(invoiceNo))
			#appendresponse.write("\n")
			#appendresponse.close()

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
	emailbody['Subject'] = "soultree daily orders sync report"
	emailbody['From'] = "soultree daily orders sync report<neeraj@paytmmloyal.com>"
	emailbody['To'] = "neeraj@paytmmloyal.com"


	msg = "We synced last one day soultree orders  <br>"
	part2 = MIMEText(msg, 'html')
	emailbody.attach(part2)
	server = smtplib.SMTP('smtp.gmail.com', 587)

	server.starttls()

	#toaddr=["neeraj@paytmmloyal.com"]
	toaddr=["ayushi@paytmmloyal.com","anoop@paytmmloyal.com"]
					
	#server.login("anoop@paytmmloyal.com", "mobianoop@1972")
	#server.login("campaignroi@paytmmloyal.com", "c@mp@ign@2018")
	server.login("neeraj@paytmmloyal.com", "qawsed@321")

	server.sendmail("soultree daily orders sync report<neeraj@paytmmloyal.com>", toaddr , emailbody.as_string())
	print("email line executed")
	server.quit()

	### move file
	# dest_folder = r'E:\pos\soultree\storeage'
	# m_fir = r'E:\pos\soultree\storeage\attachment'
	# nfilename = m_fir+'\\'+str(datetime.datetime.now().time() )+ '.csv'
	# os.rename(filename,nfilename) 
	# files = [nfilename]
	# for f in files:
	#     shutil.copy(f, dest_folder)

	#os. remove(filename)

except Exception as err:
	print('execption occured!')
	driver.close()
	driver.quit()
					
					
