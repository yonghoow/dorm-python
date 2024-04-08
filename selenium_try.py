from selenium import webdriver
from selenium.webdriver.common.by import By
import time

driver = webdriver.Edge()

driver.get('http://localhost:8088')

element = driver.find_element(By.ID, 'username')
element.send_keys('admin')
element = driver.find_element(By.ID, 'password')
element.send_keys('admin')

#submit the form
driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()


