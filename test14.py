from selenium import webdriver
from selenium.webdriver.common.alert import Alert
import time
import datetime
import openpyxl
import os
import shutil
import pandas as pd

#回数別継続

# 1.操作するブラウザを開く
# /Users/Kenta/Desktop/Selenium/chromedriver
driver = webdriver.Chrome('./Selenium/chromedriver')

# 2.操作するページを開く
driver.get('URL')

time.sleep(3)

driver.find_element_by_id('AdministratorCode').send_keys('ID')
driver.find_element_by_id('AdministratorPassword').send_keys('PASS')
time.sleep(1)
driver.find_element_by_css_selector('.btn.btn-primary.login_button').click()


driver.get('取得したいURL')
time.sleep(1)



# 今日を取得
today = datetime.datetime.today()

# 当月1日の値を出す
thismonth = datetime.datetime(today.year, today.month, 1)

# 前月末日の値を出す
lastmonth = thismonth + datetime.timedelta(days=-1)
lastmonthg = lastmonth.strftime("%Y/%m/%d")

driver.find_element_by_id('to_order_date').send_keys(lastmonthg)
time.sleep(1)
driver.find_element_by_id('submit_analyse').click()
time.sleep(10)
driver.find_element_by_id('submit_download').click()
time.sleep(5)
#csvダウンロード
Alert(driver).accept()
time.sleep(10)

workbook = openpyxl.load_workbook('CRM管理表.xlsx')
sheet = workbook["回数別継続"]

col_names = [ 'c{0:02d}'.format(i) for i in range(10) ]
  #=> ['c00', 'c01', 'c02', 'c03', 'c04', 'c05', 'c06', 'c07', 'c08', 'c09']

csv_input = pd.read_csv(filepath_or_buffer="/Users/sy/Downloads/購入回数分析_lumiere-shop.csv", encoding="ms932", sep=",", engine="python", names=col_names)

num1 = 5
num2 = 1
num3 = 1
while num2 < 50:
  try:
    if not csv_input.values[num1, 0]:
        print("ブレイクしたよ")
    elif int(csv_input.values[num1, 0]) == num2:
        print("elifに入ったよ")
        print(num2)
        #エクセルにいれる値をCSVから取得
        y = num3 + 4
        csvnum = csv_input.values[y, 1]
        print(csvnum)
        #エクセルのいれるセルの場所の値を取得
        x = str(3 + num2)
        sheetnum2 = str(sheet["A" + x].value) + "5"
        print(sheetnum2)
        sheet[sheetnum2].value  = csvnum
        num2 += 1
        num1 += 1
        num3 += 1
    else:
        print("elseに入ったよ")
        print(csv_input.values[num1, 0])
        num2 += 1
  except:
    print("最後まで終わって例外エラーになりました")
    break

workbook.save('CRM管理表.xlsx')
