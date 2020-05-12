from selenium import webdriver
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome(r'C:\Users\Admin\Desktop\chromedriver.exe')

# 100 Rows
driver.get("https://www.fieo.org/searchItcHcCode_fieo.php?stype=Like&searchStringProducts=01")

# Row count
driver.find_element_by_xpath("//*[@id='contant-contant']/div[2]/form[2]/table[2]/tbody/tr")

# Selecting element and clicking
ele = driver.find_element_by_xpath("//*[@id='contant-contant']/div[2]/form[2]/table[2]/tbody/tr[2]/td[4]/a")
ele.click()

# Switching to last window
driver.switch_to.window(driver.window_handles[-1])

# Scraping the data
info = driver.find_element_by_xpath('//*[@id="divbody"]/table/tbody/tr/td[2]/table')
info = info.replace('\n','\t')
# Converting into array
info = info.split('\t')

# Close current opened window
driver.close()

# Switch to first window
driver.switch_to.window(driver.window_handles[0])

# Creating consolidited array to paste
# array = []
array.append(info)
