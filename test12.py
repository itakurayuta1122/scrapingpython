from selenium import webdriver
import time
import datetime
import openpyxl

# エクセル自動取得RFM進捗表（パラメータ分布）

# 1.操作するブラウザを開く
# /Users/Kenta/Desktop/Selenium/chromedriver
driver = webdriver.Chrome('./Selenium/chromedriver')

# 2.操作するページを開く
driver.get('URL')

time.sleep(3)

# ログイン処理
driver.find_element_by_id('AdministratorCode').send_keys('ID')
driver.find_element_by_id('AdministratorPassword').send_keys('PASS')
time.sleep(1)
driver.find_element_by_css_selector('.btn.btn-primary.login_button').click()


driver.get('情報を淑徳したいURL')
time.sleep(1)
driver.find_element_by_id('base_date1').clear()


# 今日を取得
today = datetime.datetime.today()

# 当月1日の値を出す
thismonth = datetime.datetime(today.year, today.month, 1)

# 前月末日の値を出す
lastmonth = thismonth + datetime.timedelta(days=-1)
lastmonthg = lastmonth.strftime("%Y/%m/%d")

driver.find_element_by_id('base_date1').send_keys(lastmonthg)
time.sleep(1)
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/div[1]/form/fieldset/div/div[1]/button').click()
time.sleep(1)
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/div[1]/form/fieldset/div/div[2]/button').click()
time.sleep(1)
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/div[1]/form/fieldset/div/div[3]/button').click()
time.sleep(10)

workbook = openpyxl.load_workbook('CRM管理表.xlsx')
sheet = workbook["RFM進捗表（パラメータ分布）"]

sheet['E8'].value = int(driver.find_element_by_id('R_5_record').text.rstrip("件"))
sheet['E9'].value = int(driver.find_element_by_id('R_4_record').text.rstrip("件"))
sheet['E10'].value = int(driver.find_element_by_id('R_3_record').text.rstrip("件"))
sheet['E11'].value = int(driver.find_element_by_id('R_2_record').text.rstrip("件"))
sheet['E12'].value = int(driver.find_element_by_id('R_1_record').text.rstrip("件"))
sheet['J8'].value = int(driver.find_element_by_id('F_5_record').text.rstrip("件"))
sheet['J9'].value = int(driver.find_element_by_id('F_4_record').text.rstrip("件"))
sheet['J10'].value = int(driver.find_element_by_id('F_3_record').text.rstrip("件"))
sheet['J11'].value = int(driver.find_element_by_id('F_2_record').text.rstrip("件"))
sheet['J12'].value = int(driver.find_element_by_id('F_1_record').text.rstrip("件"))
sheet['O8'].value = int(driver.find_element_by_id('M_5_record').text.rstrip("件"))
sheet['O9'].value = int(driver.find_element_by_id('M_4_record').text.rstrip("件"))
sheet['O10'].value = int(driver.find_element_by_id('M_3_record').text.rstrip("件"))
sheet['O11'].value = int(driver.find_element_by_id('M_2_record').text.rstrip("件"))
sheet['O12'].value = int(driver.find_element_by_id('M_1_record').text.rstrip("件"))
sheet['J13'].value = driver.find_element_by_id('F_average').text.lstrip("平均累計購入回数：")
sheet['J14'].value = driver.find_element_by_id('F_max').text.lstrip("最大累計購入回数：")
sheet['O13'].value = driver.find_element_by_id('M_average').text.lstrip("平均累計購入金額：")
sheet['O14'].value = driver.find_element_by_id('M_max').text.lstrip("最大累計購入金額：")
workbook.save('CRM管理表.xlsx')
print("終わりました！")
