from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager 
from selenium.webdriver.common.by import By 
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select

import time
from bs4 import BeautifulSoup as bs
import urllib.request
import xlrd
import csv


class kush:
	def __init__(self,val):
		self.url = 'https://www.bseindia.com/corporates/ann.html'
		#self.url = 'https://www.google.com/'
		self.driver=webdriver.Chrome(ChromeDriverManager().install())
		self.delay= 2
		self.val=int(val)

	def open_page(self):
		self.driver.get(self.url)
		try:
			wait=WebDriverWait(self.driver,self.delay) 
			#print("page is ready")
		except TimeoutException:
			print("loading timeout")

	def scrap(self):
		self.driver.find_element_by_css_selector("#ddlAnnType [value='C']").click()

		y = Select(self.driver.find_element_by_xpath('//*[@id="ddlPeriod"]'))
		y.select_by_visible_text('Result')

		self.driver.find_element_by_xpath('//*[@id="txtFromDt"]').click()
		x = Select(self.driver.find_element_by_class_name('ui-datepicker-month'))
		x.select_by_visible_text('Jan')
		self.driver.find_element_by_xpath('//*[@id="ui-datepicker-div"]/table/tbody/tr[1]/td[4]/a').click()
		
		com = self.driver.find_element_by_id('scripsearchtxtbx')
		com.send_keys(self.val)
		time.sleep(4.0)
		self.driver.find_element_by_class_name('quotemenu').click()
		time.sleep(5.0)
		self.driver.find_element_by_id("btnSubmit").click()
		time.sleep(10)
		soup = bs(self.driver.page_source,'html.parser')
		print(soup)
		llist = soup.find_all('a',{'class':'tablebluelink'})
		for i in llist:
			link=i.get('href')
			print(llist)
			return link

	def quit(self):
		self.driver.close()



wb = xlrd.open_workbook("book.xlsx") 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0)
for i in range(0,1):
	val = sheet.cell_value(i, 1)
	name = sheet.cell_value(i,0)
	print(int(val), name)
	obj = kush(int(val))
	obj.open_page()
	link=obj.scrap()
	with open('ex.csv', mode='a', newline='') as file:
			writer=csv.writer(file, delimiter=',',quotechar='"',quoting=csv.QUOTE_MINIMAL)
			writer.writerow([name, val, link])
	obj.quit()


#sheet.nrows