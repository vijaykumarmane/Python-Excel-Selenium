from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import csv

driver = webdriver.Chrome(r'C:\Users\Admin\Desktop\chromedriver.exe')

array = []
HSCode = ["%.2d" % i for i in range(0,101)]

for i in range(1,3):

    driver.get("https://www.fieo.org/searchItcHcCode_fieo.php?stype=Like&searchStringProducts="+HSCode[i])
        
    # Row count
    row = driver.find_element_by_xpath("//*[@id='contant-contant']/div[2]/form[2]/table[2]/tbody/tr").size
    row = row.get('height') + 2
    print(row)
    for j in range(2, 1000):
        # Selecting element and clicking
        try:
            ele = driver.find_element_by_xpath("//*[@id='contant-contant']/div[2]/form[2]/table[2]/tbody/tr["+str(j)+"]/td[4]/a")
            ele.click()
        except:
            break
        # Switching to last window
        driver.switch_to.window(driver.window_handles[-1])
            
        # Scraping the data
        info = driver.find_element_by_xpath('//*[@id="divbody"]/table/tbody/tr/td[2]/table').text
        info = info.replace('\n','\t')
        # Converting into array
        info = info.split('\t')
        print(str(j-1)+" "+info[0])
        # Creating consolidited array to paste
        array.append(info)
        
        # Close current opened window
        driver.close()
        
        # Switch to first window
        driver.switch_to.window(driver.window_handles[0])

    # 100 Rows
    print(HSCode[i])

#file = open('g4g.csv', 'w+', newline ='\n')
#with file:
#    write = csv.writer(file)
#    write.writerows(array)
