# -*- coding: utf-8 -*-
"""
Created on Tue Feb  5 13:41:58 2019

@author: Student5
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
 

driver = webdriver.Chrome('C:/Users/student5/Desktop/ChromeDriver/chromedriver.exe') 

driver.get("https://my.palfinger.com/SiteDirectory/investments/pages/UserTasks.aspx")
driver.implicitly_wait(10)
driver.switch_to.frame(0)
window_1 = driver.window_handles[0]
driver.find_element_by_xpath("//img[@alt='edit task']").click()
driver.implicitly_wait(10)
window_2 = driver.window_handles[1]
driver.switch_to.window(window_2)
driver.find_element_by_id("btnAccept").click()
driver.find_element_by_link_text("Click to close the browser window!").click()
driver.switch_to.window(window_1)
 
# close the browser window
driver.close()
