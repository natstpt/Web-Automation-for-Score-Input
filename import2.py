import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

options = Options()
options.binary_location = 'C:/Program Files/Mozilla Firefox/firefox.exe'

driver = webdriver.Firefox(
    options=options, executable_path='geckodriver.exe')

# Assign URL
url = "http://thidamaepra.dyndns.org:81/mas.menu/"

# Opening url
driver.get(url)

# Fix wait 20sec
wait = WebDriverWait(driver, 20)

# Login
login_button1 = wait.until(EC.element_to_be_clickable(
    (By.ID, "btn-login")))
login_button1.click()

user_id = driver.find_element_by_id("user-id")
user_id.send_keys("YOUR_USER_ID")

user_pwd = driver.find_element_by_id("user-pwd")
user_pwd.send_keys("YOUR_PASSWORD")

login_button2 = wait.until(EC.element_to_be_clickable(
    (By.ID, "btn-login")))
login_button2.click()

print("Import2 Logged in")


element1 = wait.until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, "div.col-xs-6.col-sm-2.col-md-2.text-center.link-app[data-lvl='0'][data-aid='1'][data-ref='0'][data-sub='1']")))
element1.click()

element2 = wait.until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, "div.col-xs-6.col-sm-2.col-md-2.text-center.link-app[data-lvl='1'][data-aid='6'][data-ref='1'][data-sub='0']")))
element2.click()

# Wait for the options to be loaded
options = wait.until(EC.presence_of_all_elements_located(
    (By.CSS_SELECTOR, "select#sel-subject option")))

time.sleep(1)

select_subject_options = wait.until(EC.presence_of_all_elements_located(
    (By.CSS_SELECTOR, "select#sel-subject option")))

# Locate the select element
select_subject = Select(driver.find_element_by_id("sel-subject"))

""" # [รายวิชาหลัก] Select the option by value
select_subject.select_by_value("ค23102|false|false") """


# [รายวิชาเสริม] Select the option by value
select_subject.select_by_value("ค23202|false|false")

time.sleep(1)

select_class_options = wait.until(EC.presence_of_all_elements_located(
    (By.CSS_SELECTOR, "select#sel-class option")))

# Locate the select element
select_class = Select(driver.find_element_by_id("sel-class"))

# Select the option by value
select_class.select_by_value("ม.3")

time.sleep(2)

select_room_options = wait.until(EC.presence_of_all_elements_located(
    (By.CSS_SELECTOR, "select#sel-room option")))

# Locate the select element
select_room = Select(driver.find_element_by_id("sel-room"))

# Select the option by value
select_room.select_by_value("2")

time.sleep(2)

score_record = wait.until(
    EC.element_to_be_clickable((By.ID, "lb-to-score-record")))
score_record.click()

time.sleep(2)

""" # กดเมนูตรงกลาง แล้วกดเปลี่ยนเป็นกรอกคะแนนปลายภาค
# Find the element by XPath and click it
centerMenu1 = driver.find_element_by_xpath("/html/body/nav/div[2]/div/a/div/div[2]/span")
centerMenu1.click()
centerMenu2 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/div[2]/div/a[2]")))
centerMenu2.click()

print("Import2 change menu to final score on the web page successfully.") """


# read excel file
df = pd.read_excel('testExcel.xlsx', sheet_name='import2')

# store values in column B in list named "id"
id = df['id'].tolist()

# store values in column F in list named "score1"
score1 = df['score1'].tolist()

time.sleep(5)

# score1
for i in range(len(id)):
    element3 = wait.until(EC.element_to_be_clickable((By.ID, 'txt-sub-input-1-' + str(id[i]))))
    element3 = driver.find_element_by_id('txt-sub-input-1-' + str(id[i]))
    element3.send_keys(score1[i])

time.sleep(10)

# Click Show preview
element4 = wait.until(EC.element_to_be_clickable((By.ID, "lb-btn-preview")))
element4.click()

print("Import2 score1 on the web page successfully.")

""" # reset score = 0
# score1
for i in range(len(id)):
    element3 = driver.find_element_by_id('txt-sub-input-1-' + str(id[i]))
    element3.send_keys(0) 
print("Import2 reset score1 on the web page successfully.") """   