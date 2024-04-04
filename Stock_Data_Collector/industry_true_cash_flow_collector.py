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
import sys

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
    stock_index_selection.select_by_index(stock_range_index)
    time.sleep(1)
def click_export_csv():
    # 找到匯出按鈕元素
    export_button = driver.find_element(By.XPATH, '//input[@value="匯出CSV"]')
    # 點擊匯出按鈕
    export_button.click()
def top_true_volume_industry(daily_true_volume_stokc_ID, day):
    number_of_stock = []
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\產業別.xlsx').astype(int).astype(str)
    for col in industry_df.columns.values:
        n = len(industry_df.loc[industry_df[col] != '0', col].values)
        number_of_stock.append([col, n])
    number_df = pd.DataFrame(number_of_stock, columns=['industry', 'number'])
    all = []
    for id in daily_true_volume_stokc_ID:
        stock_ind = []
        for col in industry_df.columns.values:
            # print(id)
            if str(id) in industry_df[col].values:
                stock_ind.append(col)
                all.append([col, id, 1])
    all_df = pd.DataFrame(all, columns=['industry', 'ID', 'count'])
    print(all_df.loc[all_df['industry']=='穿戴式裝置'])
    # print(all_df.head())
    top_volume_industry = all_df.groupby(by='industry').sum().sort_values(by='count',ascending=False)
    print(top_volume_industry.loc['穿戴式裝置', 'count'])
    top_volume_industry['all number'] = 1
    for industry in top_volume_industry.index.values:
        top_volume_industry.loc[industry, 'all number'] = number_df.loc[number_df['industry']==industry, 'number'].values
        top_volume_industry.loc[industry, 'percentage'] = round(100*top_volume_industry.loc[industry, 'count']/top_volume_industry.loc[industry, 'all number'],1)
    
    print(top_volume_industry.loc['穿戴式裝置', 'percentage'])
    top_volume_industry = top_volume_industry.transpose()
    daily_top_volume_percentage = pd.DataFrame(top_volume_industry.loc['percentage'].values,columns=[f'2024/{str(day)}'], index=top_volume_industry.columns.values).transpose()
    history = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/100產業實質成交值排行.xlsx', header=0, index_col=0)
    history.index = history.index.astype(str)
    try:
        history.drop(str(day), inplace=True)
        print(f'already update, rewrite today top business volume industry.')
    except Exception as e:
        pass
    h = pd.concat([history, daily_top_volume_percentage], axis =0)
    
    h.sort_index(ascending = False, inplace=True)
    h.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/100產業實質成交值排行.xlsx')
service = ChromeService(executable_path=ChromeDriverManager().install())
# 關閉通知提醒
chrome_options = webdriver.ChromeOptions()
prefs = {"profile.default_content_setting_values.notifications" : 2}
chrome_options.add_experimental_option("prefs",prefs)
# chrome_options.add_argument('--headless')
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

# 開啟瀏覽器
url = 'https://goodinfo.tw/tw2/StockList.asp?RPT_TIME=&MARKET_CAT=\
%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E7%8F%BE%E8%82%A1\
%E7%95%B6%E6%B2%96%E5%BC%B5%E6%95%B8+%28%E7%95%B6%E6%97%A5%29%40%40%E\
7%8F%BE%E8%82%A1%E7%95%B6%E6%B2%96%E5%BC%B5%E6%95%B8%40%40%E7%95%B6%E6%97%A5'
for i in range(4): # 5:一周
    #當沖日期 1: 今日 2: 昨日
    choose_date = i+1
    driver = webdriver.Chrome(service=service, options=chrome_options)
    try:
        driver.get(url)
    except Exception as e:
        print(e)
        driver.close()
        driver.quit()
    # ok = input('請確認網頁已經開啟，按下Enter繼續')
    for date_index in [choose_date]:
        select_date(date_index)
        if date_index == 1:
            print('今日csv')
        elif date_index == 2:
            print('昨日csv')
        for stock_range_index in range(6):
            select_stock_index(stock_range_index)
            time.sleep(2.5)
            click_export_csv()
            print(f'已匯出csv {stock_range_index+1}')
            
        time.sleep(3)
    driver.close()


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
    # 讀取每個 CSV 檔案並進行處理
    for j, file_name in enumerate(file_names):
        file_path = os.path.join(folder_path, file_name)
        print('處理檔案：', file_name)
        read_df = pd.read_csv(file_path)
        file_date = read_df['當沖日期'].values[0].replace('/','_')
        print(read_df['當沖日期'].values[0])
        print(file_date)
        save_file_name = file_date +'.csv'
        save_path = 'C:/Users/User/Desktop/實質成交量資料'
        df = read_df[['排名','代號','名稱','當沖日期','當沖買進額(百萬)','成交額(百萬)']].copy()
        df['代號'] = df['代號'].map(lambda x: x.replace('=','').replace('"',''))
        df[['成交額(百萬)', '當沖買進額(百萬)']] = df[['成交額(百萬)', '當沖買進額(百萬)']].astype(str)
        if df['成交額(百萬)'].dtype == 'object':
            df['成交額(百萬)'] = df['成交額(百萬)'].map(lambda x: x.replace(',',''))
        if df['當沖買進額(百萬)'].dtype == 'object':
            df['當沖買進額(百萬)'] = df['當沖買進額(百萬)'].map(lambda x: x.replace(',',''))
        df['實質成交量(百萬)'] = df['成交額(百萬)'].astype(float) - df['當沖買進額(百萬)'].astype(float)
        # write if file not in save_path
        if j == 0:
            if save_file_name in os.listdir(save_path):
                print('already have today data, remove '+os.path.join(save_path, save_file_name))
                os.remove(os.path.join(save_path, save_file_name))

        # sys.exit()
        # 第一次寫入
        if save_file_name not in os.listdir(save_path):
            df = df.loc[list(map(lambda x:len(x) == 4, df['代號'].values))]
            df.to_csv(os.path.join(save_path, save_file_name), index=False, encoding='utf-8-sig')
        else:
            # print(df['代號'].values[0], len(df['代號'].values[0]))
            the_date_df = pd.read_csv(os.path.join(save_path, save_file_name))
            df = df.loc[list(map(lambda x:len(x) == 4, df['代號']))]
            the_date_df = pd.concat([the_date_df,df])
            the_date_df.to_csv(os.path.join(save_path, save_file_name), index=False, encoding='utf-8-sig')
        # os.remove(file_path)
    for remove_file_name in file_names:
        file_path = os.path.join(folder_path, remove_file_name)
        os.remove(file_path)

    # 資料夾路徑
    folder_path = 'C:/Users/User/Desktop/實質成交量資料'

    # 獲取資料夾中所有的檔案名稱
    # file_names = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
    file_names = [save_file_name]
    for file_name in file_names:
        date = file_name.replace('.csv','')
        print(date)
        daily_true_volume_df = pd.read_csv('C:/Users/User/Desktop/實質成交量資料/'+file_name)
        daily_true_volume_df = daily_true_volume_df.sort_values(by='實質成交量(百萬)', ascending=False)
        # print(daily_true_volume_df.head())
        daily_true_volume_df.set_index('代號', inplace=True)
        daily_true_volume_df = daily_true_volume_df.iloc[:340]
        daily_true_volume_stokc_ID = daily_true_volume_df.index.values
        top_true_volume_industry(daily_true_volume_stokc_ID, date.replace('_','/'))

