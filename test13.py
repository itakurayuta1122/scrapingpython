from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
import time
import datetime
import openpyxl
import os
import shutil
import pandas as pd

#CPM分析

workbook = openpyxl.load_workbook('CRM管理表.xlsx')
sheet = workbook["ルミエールCPM"]

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
time.sleep(15)

sheet["D3"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[2]/td[14]/span').text
sheet["D4"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[3]/td[20]/span').text
sheet["D7"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[4]/td[14]/span').text
sheet["D8"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[5]/td[20]/span').text
sheet["D11"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[6]/td[14]/span').text
sheet["D12"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[7]/td[20]/span').text
sheet["D15"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[8]/td[14]/span').text
sheet["D16"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[9]/td[20]/span').text
sheet["D19"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[10]/td[14]/span').text
sheet["D20"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[11]/td[20]/span').text

sheet["E3"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[2]/td[12]/span').text
sheet["E4"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[3]/td[17]/span').text
sheet["E7"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[4]/td[12]/span').text
sheet["E8"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[5]/td[17]/span').text
sheet["E11"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[6]/td[12]/span').text
sheet["E12"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[7]/td[17]/span').text
sheet["E15"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[8]/td[12]/span').text
sheet["E16"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[9]/td[17]/span').text
sheet["E19"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[10]/td[12]/span').text
sheet["E20"].value = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[3]/div/div/div/table[3]/tbody/tr[11]/td[17]/span').text
print("終わりました")

workbook.save('CRM管理表.xlsx')
