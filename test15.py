from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
import time
import datetime
import openpyxl
import os
import shutil
import pandas as pd

# 今日を取得
today = datetime.datetime.today()

# 当月1日の値を出す
thismonth = datetime.datetime(today.year, today.month, 1)

# 前月末日の値を出す
lastmonth = thismonth + datetime.timedelta(days=-1)
lastmonthg = lastmonth.strftime("%Y/%m/%d")

lastmonfir = datetime.datetime(lastmonth.year, lastmonth.month, 1)
lastmonfirst = lastmonfir.strftime("%Y/%m/%d")


# 1.操作するブラウザを開く
# /Users/Kenta/Desktop/Selenium/chromedriver
driver = webdriver.Chrome('./Selenium/chromedriver')

# 2.操作するページを開く
driver.get('https://secure.uchideno-kozuchi.com/crm/lumiere-shop/admin/administrators/login')

time.sleep(3)

driver.find_element_by_id('AdministratorCode').send_keys('lumiere-shop')
driver.find_element_by_id('AdministratorPassword').send_keys('lumiere-shop_admin')
time.sleep(1)
driver.find_element_by_css_selector('.btn.btn-primary.login_button').click()

driver.get('https://secure.uchideno-kozuchi.com/crm/lumiere-shop/admin/analyses/order_count')
time.sleep(1)

driver.find_element_by_id('from_order_date').send_keys(lastmonfirst)
driver.find_element_by_id('to_order_date').send_keys(lastmonthg)

driver.find_element_by_css_selector('.icon-off').click()
time.sleep(2)
Select(driver.find_element_by_css_selector('.input-xlarge.with-button-extra.focused_segment_selector_select')).select_by_value("40")
time.sleep(1)
driver.find_element_by_id('submit_analyse').click()
time.sleep(10)
driver.find_element_by_id('submit_download').click()
time.sleep(5)
#当月csvダウンロード
Alert(driver).accept()
time.sleep(10)
driver.find_element_by_id('from_order_date').clear()
time.sleep(1)
driver.find_element_by_id('submit_analyse').click()
time.sleep(10)
driver.find_element_by_id('submit_download').click()
time.sleep(5)
#全期間csvダウンロード
Alert(driver).accept()
time.sleep(10)

driver.get('https://secure.uchideno-kozuchi.com/crm/lumiere-shop/admin/analyses/order_count')
time.sleep(5)

driver.find_element_by_id('from_order_date').send_keys(lastmonfirst)
driver.find_element_by_id('to_order_date').send_keys(lastmonthg)

time.sleep(2)
Select(driver.find_element_by_css_selector('.input-xlarge.with-button-extra.focused_segment_selector_select')).select_by_value("39")
time.sleep(1)
driver.find_element_by_id('submit_analyse').click()
time.sleep(10)
driver.find_element_by_id('submit_download').click()
time.sleep(5)
#当月csvダウンロード
Alert(driver).accept()
time.sleep(10)
driver.find_element_by_id('from_order_date').clear()
time.sleep(1)
driver.find_element_by_id('submit_analyse').click()
time.sleep(10)
driver.find_element_by_id('submit_download').click()
time.sleep(5)
#全期間csvダウンロード
Alert(driver).accept()
time.sleep(10)

driver.get('https://secure.uchideno-kozuchi.com/crm/lumiere-shop/admin/analyses/order_count')
time.sleep(5)

driver.find_element_by_id('from_order_date').send_keys(lastmonfirst)
driver.find_element_by_id('to_order_date').send_keys(lastmonthg)

time.sleep(2)
Select(driver.find_element_by_css_selector('.input-xlarge.with-button-extra.focused_segment_selector_select')).select_by_value("38")
time.sleep(1)
driver.find_element_by_id('submit_analyse').click()
time.sleep(10)
driver.find_element_by_id('submit_download').click()
time.sleep(5)
#当月csvダウンロード
Alert(driver).accept()
time.sleep(10)
driver.find_element_by_id('from_order_date').clear()
time.sleep(1)
driver.find_element_by_id('submit_analyse').click()
time.sleep(10)
driver.find_element_by_id('submit_download').click()
time.sleep(5)
#全期間csvダウンロード
Alert(driver).accept()
time.sleep(10)

driver.get('https://secure.uchideno-kozuchi.com/crm/lumiere-shop/admin/analyses/order_count')
time.sleep(5)

driver.find_element_by_id('from_order_date').send_keys(lastmonfirst)
driver.find_element_by_id('to_order_date').send_keys(lastmonthg)

time.sleep(2)
Select(driver.find_element_by_css_selector('.input-xlarge.with-button-extra.focused_segment_selector_select')).select_by_value("37")
time.sleep(1)
driver.find_element_by_id('submit_analyse').click()
time.sleep(10)
driver.find_element_by_id('submit_download').click()
time.sleep(5)
#当月csvダウンロード
Alert(driver).accept()
time.sleep(10)
driver.find_element_by_id('from_order_date').clear()
time.sleep(1)
driver.find_element_by_id('submit_analyse').click()
time.sleep(10)
driver.find_element_by_id('submit_download').click()
time.sleep(5)
#全期間csvダウンロード
Alert(driver).accept()
time.sleep(10)


workbook = openpyxl.load_workbook('CRM管理表.xlsx')
sheet = workbook["商品別転換率"]

col_names = [ 'c{0:02d}'.format(i) for i in range(10) ]
  #=> ['c00', 'c01', 'c02', 'c03', 'c04', 'c05', 'c06', 'c07', 'c08', 'c09']
col_names2 = [ 'c{0:02d}'.format(i) for i in range(10) ]

csv_input = pd.read_csv(filepath_or_buffer="/Users/sy/Downloads/購入回数分析_lumiere-shop.csv", encoding="ms932", sep=",", engine="python", names=col_names)

csv_input2 = pd.read_csv(filepath_or_buffer="/Users/sy/Downloads/購入回数分析_lumiere-shop (1).csv", encoding="ms932", sep=",", engine="python", names=col_names2)

try:
  if(int(csv_input.values[6, 0]) == 2):
    print(csv_input.values[6, 1])
    sheet["D8"].value = csv_input.values[6, 1]

except:
  print("二回め購入はない")
  sheet["D8"].value = 0

try:
  x = csv_input['c01'].astype(str)
  y = x.drop([0,1,2,3,4]).astype(int)
  print(y.sum())
  sheet["C8"].value = y.sum()
except:
  print("全然購入はない")
  sheet["C8"].value = 0


try:
  if(int(csv_input2.values[6, 0]) == 2):
    print(csv_input2.values[6, 1])
    sheet["G8"].value = csv_input2.values[6, 1]
except:
  print("二回め購入はない")
  sheet["G8"].value = 0

try:
  x = csv_input2['c01'].astype(str)
  y = x.drop([0,1,2,3,4]).astype(int)
  print(y.sum())
  sheet["F8"].value = y.sum()
except:
  print("全然購入はない")
  sheet["F8"].value = 0

csv_input = pd.read_csv(filepath_or_buffer="/Users/sy/Downloads/購入回数分析_lumiere-shop (2).csv", encoding="ms932", sep=",", engine="python", names=col_names)

csv_input2 = pd.read_csv(filepath_or_buffer="/Users/sy/Downloads/購入回数分析_lumiere-shop (3).csv", encoding="ms932", sep=",", engine="python", names=col_names2)

try:
  if(int(csv_input.values[6, 0]) == 2):
    print(csv_input.values[6, 1])
    sheet["D6"].value = csv_input.values[6, 1]

except:
  print("二回め購入はない")
  sheet["D6"].value = 0

try:
  x = csv_input['c01'].astype(str)
  y = x.drop([0,1,2,3,4]).astype(int)
  print(y.sum())
  sheet["C6"].value = y.sum()
except:
  print("全然購入はない")
  sheet["C6"].value = 0


try:
  if(int(csv_input2.values[6, 0]) == 2):
    print(csv_input2.values[6, 1])
    sheet["G6"].value = csv_input2.values[6, 1]
except:
  print("二回め購入はない")
  sheet["G6"].value = 0

try:
  x = csv_input2['c01'].astype(str)
  y = x.drop([0,1,2,3,4]).astype(int)
  print(y.sum())
  sheet["F6"].value = y.sum()
except:
  print("全然購入はない")
  sheet["F6"].value = 0




csv_input = pd.read_csv(filepath_or_buffer="/Users/sy/Downloads/購入回数分析_lumiere-shop (4).csv", encoding="ms932", sep=",", engine="python", names=col_names)

csv_input2 = pd.read_csv(filepath_or_buffer="/Users/sy/Downloads/購入回数分析_lumiere-shop (5).csv", encoding="ms932", sep=",", engine="python", names=col_names2)

try:
  if(int(csv_input.values[6, 0]) == 2):
    print(csv_input.values[6, 1])
    sheet["D7"].value = csv_input.values[6, 1]

except:
  print("二回め購入はない")
  sheet["D7"].value = 0

try:
  x = csv_input['c01'].astype(str)
  y = x.drop([0,1,2,3,4]).astype(int)
  print(y.sum())
  sheet["C7"].value = y.sum()
except:
  print("全然購入はない")
  sheet["C7"].value = 0


try:
  if(int(csv_input2.values[6, 0]) == 2):
    print(csv_input2.values[6, 1])
    sheet["G7"].value = csv_input2.values[6, 1]
except:
  print("二回め購入はない")
  sheet["G7"].value = 0

try:
  x = csv_input2['c01'].astype(str)
  y = x.drop([0,1,2,3,4]).astype(int)
  print(y.sum())
  sheet["F7"].value = y.sum()
except:
  print("全然購入はない")
  sheet["F7"].value = 0

csv_input = pd.read_csv(filepath_or_buffer="/Users/sy/Downloads/購入回数分析_lumiere-shop (6).csv", encoding="ms932", sep=",", engine="python", names=col_names)

csv_input2 = pd.read_csv(filepath_or_buffer="/Users/sy/Downloads/購入回数分析_lumiere-shop (7).csv", encoding="ms932", sep=",", engine="python", names=col_names2)

try:
  if(int(csv_input.values[6, 0]) == 2):
    print(csv_input.values[6, 1])
    sheet["D5"].value = csv_input.values[6, 1]

except:
  print("二回め購入はない")
  sheet["D5"].value = 0

try:
  x = csv_input['c01'].astype(str)
  y = x.drop([0,1,2,3,4]).astype(int)
  print(y.sum())
  sheet["C5"].value = y.sum()
except:
  print("全然購入はない")
  sheet["C5"].value = 0


try:
  if(int(csv_input2.values[6, 0]) == 2):
    print(csv_input2.values[6, 1])
    sheet["G5"].value = csv_input2.values[6, 1]
except:
  print("二回め購入はない")
  sheet["G5"].value = 0

try:
  x = csv_input2['c01'].astype(str)
  y = x.drop([0,1,2,3,4]).astype(int)
  print(y.sum())
  sheet["F5"].value = y.sum()
except:
  print("全然購入はない")
  sheet["F5"].value = 0

workbook.save('CRM管理表.xlsx')


