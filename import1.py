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

print("Import1 Logged in")

element1 = wait.until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, "div.col-xs-6.col-sm-2.col-md-2.text-center.link-app[data-lvl='0'][data-aid='1'][data-ref='0'][data-sub='1']")))
element1.click()

element2 = wait.until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, "div.col-xs-6.col-sm-2.col-md-2.text-center.link-app[data-lvl='1'][data-aid='5'][data-ref='1'][data-sub='0']")))
element2.click()

# Wait for the options to be loaded
options = wait.until(EC.presence_of_all_elements_located(
    (By.CSS_SELECTOR, "select#sel-subject option")))

time.sleep(1)

select_subject_options = wait.until(EC.presence_of_all_elements_located(
    (By.CSS_SELECTOR, "select#sel-subject option")))

# Locate the select element
select_subject = Select(driver.find_element_by_id("sel-subject"))

""" # [รายวิชาหลัก] Select the option by its index 
select_subject.select_by_value("ค23102|false|false|3|คณิตศาสตร์ 6|คณิตศาสตร์ 6|1.5") """

# [รายวิชาเสริม] Select the option by its index 
select_subject.select_by_value("ค23202|false|false|3|คณิตศาสตร์ (เสริม) 6|คณิตศาสตร์ (เสริม) 6|1")

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

select_prog_options = wait.until(EC.presence_of_all_elements_located(
    (By.CSS_SELECTOR, "select#sel-prog option")))

# Locate the select element
select_prog = Select(driver.find_element_by_id("sel-prog"))

# Select the option by value
select_prog.select_by_value("104|MLP")

score_record = wait.until(
    EC.element_to_be_clickable((By.ID, "lb-to-score-record")))
score_record.click()

time.sleep(2)

# read excel file
df = pd.read_excel('testExcel.xlsx', sheet_name='import1')

# store values in column B in list named "id"
id = df['id'].tolist()

# store values in column F in list named "score1"
score1 = df['score1'].tolist()

# store values in column G in list named "score2"
score2 = df['score2'].tolist()

# store values in column H in list named "score3"
score3 = df['score3'].tolist()

# store values in column H in list named "score3"
score4 = df['score4'].tolist()

""" # store values in column H in list named "score3"
score5 = df['score5'].tolist() """

# store values in column H in list named "score3"
score6 = df['score6'].tolist()

time.sleep(5)

# score1
for i in range(len(id)):
    element3 = wait.until(EC.element_to_be_clickable((By.ID, 'txt-sub-input-p1-' + str(id[i]))))
    element3 = driver.find_element_by_id('txt-sub-input-p1-' + str(id[i]))
    element3.send_keys(score1[i])

# score2
for i in range(len(id)):
    element4 = wait.until(EC.element_to_be_clickable((By.ID, 'txt-sub-input-p2-' + str(id[i]))))
    element4 = driver.find_element_by_id('txt-sub-input-p2-' + str(id[i]))
    element4.send_keys(score2[i])

# score3
for i in range(len(id)):
    element5 = wait.until(EC.element_to_be_clickable((By.ID, 'txt-sub-input-p2-' + str(id[i]))))
    element5 = driver.find_element_by_id('txt-sub-input-p3-' + str(id[i]))
    element5.send_keys(score3[i])

# score4
for i in range(len(id)):
    element6 = wait.until(EC.element_to_be_clickable((By.ID, 'txt-sub-input-p2-' + str(id[i]))))
    element6 = driver.find_element_by_id('txt-sub-input-p4-' + str(id[i]))
    element6.send_keys(score4[i])

""" # score5
for i in range(len(id)):
    element7 = wait.until(EC.element_to_be_clickable((By.ID, 'txt-sub-input-p2-' + str(id[i]))))
    element7 = driver.find_element_by_id('txt-sub-input-p5-' + str(id[i]))
    element7.send_keys(score5[i])
 """
print("Import1 score1 score2 score3 score4 on the web page successfully.")

# กดเมนูตรงกลาง แล้วกดเปลี่ยนเป็นกรอกคะแนนปลายภาค
# Find the element by XPath and click it
centerMenu1 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/nav/div[2]/div/a/div/div[2]/span")))
centerMenu1.click()
centerMenu2 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/div[2]/div/a[2]")))
centerMenu2.click()

print("Import1 change menu to final score on the web page successfully.")

time.sleep(10)

# score4
for i in range(len(id)):
    element8 = wait.until(EC.element_to_be_clickable((By.ID, 'txt-sub-input-p1-' + str(id[i]))))
    element8 = driver.find_element_by_id('txt-sub-input-p1-' + str(id[i]))
    element8.send_keys(score6[i])

time.sleep(5)

# Click Show preview
element9 = wait.until(EC.element_to_be_clickable((By.ID, "lb-btn-preview")))
element9.click()

print("Import1 score6 on the web page successfully.")


""" # reset score = 0
# score1
for i in range(len(id)):
    element3 = driver.find_element_by_id('txt-sub-input-p1-' + str(id[i]))
    element3.send_keys(0)

# score2
for i in range(len(id)):
    element4 = driver.find_element_by_id('txt-sub-input-p2-' + str(id[i]))
    element4.send_keys(0)

# score3
for i in range(len(id)):
    element5 = driver.find_element_by_id('txt-sub-input-p3-' + str(id[i]))
    element5.send_keys(0) 

print("Import1 reset score1 score2 score3 on the web page successfully.")
    """