from selenium import webdriver
from selenium.webdriver.common.by import By

def get_account_data():

    driver = webdriver.Chrome()

    driver.get("https://www.uccu.com")
    driver.implicitly_wait(5)

    userid_box = driver.find_element(by=By.NAME, value="user_id")
    password_box = driver.find_element(by=By.NAME, value="password")
    sign_in_button = driver.find_element(by=By.CSS_SELECTOR, value="button")

    userid_box.send_keys("1095729")
    password_box.send_keys("M.,mlkjpou75")
    sign_in_button.click()
    driver.implicitly_wait(5)


    title = driver.title
    assert title == "Utah Community Federal Credit Union: Homeasdf"



    driver.quit()