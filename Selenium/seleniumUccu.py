from selenium import webdriver
from selenium.webdriver.common.by import By

# def get_account_data():

driver = webdriver.Chrome()
driver.get("https://www.uccu.com")
driver.implicitly_wait(5)
userid_box = driver.find_element(by=By.ID, value="QuickLoginHeader_user_id")
password_box = driver.find_element(by=By.ID, value="QuickLoginHeader_password")
sign_in_form = driver.find_element(by=By.ID, value="Q2OnlineLogin")
# sign_in_button = driver.find_element_by_css_selector(".Button")
userid_box.send_keys("1095729")
password_box.send_keys("M.,mlkjpou75")
sign_in_form.submit()
driver.implicitly_wait(5)
title = driver.title
assert title == "Utah Community Federal Credit Union: Homeasdf"



    # driver.quit()