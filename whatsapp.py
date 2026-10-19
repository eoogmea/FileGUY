from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys

#enter perfect name as in your contact (it is case sensitive)
name= "Babbi"
#replace /chromedriver with loaction of chromedriver on your computer
options = webdriver.ChromeOptions();
options.add_argument('--user-data-dir=./User_Data')
driver = webdriver.Chrome(chrome_options=options)
driver.get('https://web.whatsapp.com/')
time.sleep(10)
elem = driver.find_element_by_xpath("//input[@title='Cerca o inizia una nuova chat']")
elem.click()
elem.send_keys(name)
time.sleep(10)
driver.find_element_by_xpath("//*[@title='"+name+"']").click();
time.sleep(5)
elem = driver.find_element_by_xpath("//div[@data-tab='1']")
elem.click()
for i in range(20):
 elem.send_keys("hello this is from python " + Keys.ENTER)
 time.sleep(1)