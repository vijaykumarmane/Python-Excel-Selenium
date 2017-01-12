from selenium import webdriver

# Install geckodriver.exe appropriate to OS bit and add environment
# variable PATH to driver to work on new updated firefox version and same goes
# to chromedriver; windows 32 and 64 has same chromedriver 32 bit

driver = webdriver.Chrome(r'C:\Users\vijaykumar.mane\Desktop\chromedriver.exe')
# driver = webdriver.Chrome() this way chrome not works

# driver = webdriver.Firefox(r'C:\Users\vijaykumar.mane\Desktop\geckodriver.exe') this way its not works but chrome works

driver.get("http://www.python.org")

driver.close()  # Issue to close Chrome browser.
