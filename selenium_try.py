# import required modules
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

#add headless options
#options = Options()
#options.add_argument('headless')

#create object with headless options
#driver = webdriver.Edge(options = options)
driver = webdriver.Edge()

#open browser and navigate to webpage
#driver.get('http://localhost:8088')
#driver.get('https://www.duckduckgo.com')
driver.get('https://www.python.org')

#search elements
#element = driver.find_element(By.CLASS_NAME, 'searchbox_input__bEGm3')
#element.send_keys('real python')
#element.submit()

#element = driver.find_element(By.ID, 'username')
#element.send_keys('admin')
#element.submit()
#element = driver.find_element(By.ID, 'password')
#element.send_keys('admin')
#element.submit()

#submit the form
#driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

#prints contents of <title></title>
print(driver.title)

#find DOM element by name for search bar
search_bar = driver.find_element(By.NAME, 'q')

#clear contents of search bar
search_bar.clear()

#enter a string as its value using .send_keys()
search_bar.send_keys('getting started with python')

#emulate the press of Return key using Keys.RETURN
search_bar.send_keys(Keys.RETURN)

#actions trigger a new url. Print current url
print(driver.current_url)

#close browser object
#driver.close()

