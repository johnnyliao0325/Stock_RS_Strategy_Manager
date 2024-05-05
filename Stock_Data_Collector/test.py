from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
# import pyautogui
import time
import requests
import os
import pandas as pd
from io import StringIO
import os
import pandas as pd
# 選日期
#當沖日期 1: 今日 2: 昨日
choose_date = 1
def select_date(date_index):
    select_date_element = driver.find_element(By.ID, 'selRPT_TIME')
    date_selection = Select(select_date_element)
    # print(len(date_selection.options))
    date_selection.select_by_index(date_index)
    time.sleep(1)
def select_stock_index(stock_range_index):
    select_stock_index_element = driver.find_element(By.ID, 'selRANK')
    stock_index_selection = Select(select_stock_index_element)
    # print(len(stock_index_selection.options))
    # stock_index_selection.select_by_index(stock_range_index)
    stock_index_selection.select_by_value(str(stock_range_index))
    time.sleep(1)
def select_info(select_info_index_element):
    select_stock_index_element = driver.find_element(By.ID, 'selSHEET')
    stock_index_selection = Select(select_info_index_element)
    stock_index_selection.select_by_value(str(select_info_index_element))
def select_info2(select_info_index_element2):
    select_stock_index_element = driver.find_element(By.ID, 'selSHEET2')
    stock_index_selection = Select(select_info_index_element2)
    stock_index_selection.select_by_value(str(select_info_index_element2))
def click_export_csv():
    # 找到匯出按鈕元素
    export_button = driver.find_element(By.XPATH, '//input[@value="匯出CSV"]')
    # 點擊匯出按鈕
    export_button.click()
service = ChromeService(executable_path=ChromeDriverManager().install())
# 關閉通知提醒
chrome_options = webdriver.ChromeOptions()
prefs = {"profile.default_content_setting_values.notifications" : 2}
chrome_options.add_experimental_option("prefs",prefs)
# chrome_options.add_argument('--headless')
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

# 開啟瀏覽器
url = 'https://goodinfo.tw/tw2/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=\
%E5%96%AE%E5%AD%A3%E5%AD%98%E8%B2%A8%E9%80%B1%E8%BD%89%E7%8E%87%E6%9C%80%E9%AB%98%40%40%E5%AD%98%E8%B2%A8%E9%80%B1%E8\
%BD%89%E7%8E%87%40%40%E5%96%AE%E5%AD%A3%E6%9C%80%E9%AB%98'

driver = webdriver.Chrome(service=service, options=chrome_options)
try:
    driver.get(url)
except Exception as e:
    print(e)
    driver.close()
    driver.quit()
# ok = input('請確認網頁已經開啟，按下Enter繼續')
# for date_index in [choose_date]:
#     select_date(date_index)
#     if date_index == 1:
#         print('今日csv')
#     elif date_index == 2:
#         print('昨日csv')
#     for stock_range_index in range(6):
#         select_stock_index(stock_range_index)
#         time.sleep(2)
#         click_export_csv()
#         print(f'已匯出csv {stock_range_index+1}')
#     time.sleep(3)
# driver.close()
# driver.quit()
select_info(30)
time.sleep(5)
for stock_range_index in range(6):
    select_stock_index(stock_range_index)
    time.sleep(2)
    click_export_csv()
    print(f'已匯出csv {stock_range_index+1}')
time.sleep(3)
driver.close()
driver.quit()
# 資料夾路徑
folder_path = 'C:/Users/User/Downloads/'

# 獲取資料夾中所有的檔案名稱
# file_names = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
file_names = []
for i in range(6):
    if i == 0:
        file_names.append(f'StockList.csv')
    else:
        file_names.append(f'StockList ({i}).csv')
all_df = pd.DataFrame()
for file_name in file_names:
    file_path = os.path.join(folder_path, file_name)
    print('處理檔案：', file_name)
    read_df = pd.read_csv(file_path)
    df = df.loc[list(map(lambda x:len(x) == 4, df['代號'].values))]
    all_df = pd.concat([all_df, df])
    os.remove(file_path)
all_df.to_csv('C:/Users/User/Downloads/StockList.csv', index=False)
