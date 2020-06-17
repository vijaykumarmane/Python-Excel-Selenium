import xlwings as xw
from selenium import webdriver
import re
import time
from fuzzywuzzy import fuzz
import os

# Open saved workbook
wb = xw.Book(os.path.expanduser("~")+"\\Desktop\\Contact Details.xlsx")
sht = wb.sheets[0]

driver = webdriver.Chrome(os.path.expanduser("~")+"\\Desktop\\chromedriver.exe")

def ecosiaEmail(companyName):
	url = "https://www.ecosia.org/search?q=Zauba+corp+" + companyName.replace("& ","")
	url = url.replace(' ','+')
	time.sleep(3)
	driver.get(url)
	time.sleep(2)
	email, Ratio = "", ""
	try:
		link = driver.find_element_by_partial_link_text("https://www.zaubacorp.com/company/")
		time.sleep(1)
		comp = re.search('https://www.zaubacorp.com/company/(.+)/', link.text).group(1).replace("-"," ")
		Ratio = fuzz.ratio(companyName.lower().replace("& ","").replace("pvt.","private").replace('ltd.','limited'),comp.lower())
		if Ratio > 85:
			link.click()
			time.sleep(3)
			ele = driver.find_element_by_xpath('/html/body')
			email = re.search('Email ID: ([\S+@\S+]+)', ele.text)
			email = email.group(1)
			return email, Ratio
		else:
			return "", ""
	except:
		return "", ""

def googleEmail(companyName):

	url = "https://www.google.com/search?q=zaubacorp.com " + companyName.replace("& ","")
	url = url.replace(' ','+')
	driver.get(url)
	time.sleep(3)

	try:
		email, Ratio = "", ""
		link = driver.find_element_by_tag_name('cite')
		time.sleep(1)
		Ratio = 0
		if "www.zaubacorp.com › company ›" in link.text:
			link.click()
			time.sleep(3)
			ele = driver.find_element_by_xpath('/html/body')
			email = re.search('Email ID: ([\S+@\S+]+)', ele.text)
			email = email.group(1)
			return email, Ratio
		else:
			return "", ""
	except:
		return "", ""

i = 2
while sht.range("B" + str(i)).value != None:
	companyName = sht.range("B" + str(i)).value
	prevName = sht.range("B" + str(i-1)).value
	if companyName != prevName:
		output = ecosiaEmail(companyName)
		sht.range("C" + str(i)).value = output
		if not "@" in str(output[0]):
			output = googleEmail(companyName)
			sht.range("C" + str(i)).value = output
	else:
		sht.range("C"+str(i)+":"+"D"+str(i)).value = sht.range("C"+str(i-1)+":"+"D"+str(i-1)).value
	i+=1

wb.save()
driver.close()

print('Done')

