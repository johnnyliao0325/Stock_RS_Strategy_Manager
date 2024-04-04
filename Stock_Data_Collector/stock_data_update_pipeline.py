import requests
import pandas as pd
import datetime
import warnings
import sys
sys.path.append('C:/Users/User/Desktop/StockInfoHub')
from Shared_Modules.shared_functions import *
from Shared_Modules.shared_variables import *
from stock_data_collector_package.stock_data_collector_utils import *

# Suppress all warnings
warnings.filterwarnings("ignore")
class bcolors:
    OK = '\033[92m' #GREEN
    WARNING = '\033[93m' #YELLOW
    FAIL = '\033[91m' #RED
    RESET = '\033[0m' #RESET COLOR
# ============上市股票df============
url = "https://isin.twse.com.tw/isin/class_main.jsp?owncode=&stockname=&isincode=&market=1&issuetype=1&industry_code=&Page=1&chklike=Y"
response = requests.get(url)
listed = pd.read_html(response.text)[0]
listed.columns = listed.iloc[0,:]
listed = listed[["有價證券代號","有價證券名稱","市場別","產業別","公開發行/上市(櫃)/發行日"]]
listed = listed.iloc[1:]
# ============上櫃股票df============
urlTWO = "https://isin.twse.com.tw/isin/class_main.jsp?owncode=&stockname=&isincode=&market=2&issuetype=&industry_code=&Page=1&chklike=Y"
response = requests.get(urlTWO)
listedTWO = pd.read_html(response.text)[0]
listedTWO.columns = listedTWO.iloc[0,:]
listedTWO = listedTWO.loc[listedTWO['有價證券別'] == '股票']
listedTWO = listedTWO[["有價證券代號","有價證券名稱","市場別","產業別","公開發行/上市(櫃)/發行日"]]
# ============上市股票代號+.TW============
stock_1 = listed["有價證券代號"]
stock_num = stock_1.apply(lambda x: str(x) + ".TW")
taiwan_0050 = pd.Series(['0050.TW'], index=[0])
taiwan_weighted = pd.Series(['^TWII'], index=[0])
# print(stock_num)
# ============上櫃股票代號+.TWO============
stock_2 = listedTWO["有價證券代號"]
stock_num2 = stock_2.apply(lambda x: str(x) + ".TWO")
# print(stock_num2)
# ============concat全部股票代號============
stock_num = pd.concat([taiwan_weighted, taiwan_0050, stock_num, stock_num2], ignore_index=True)
# print(stock_num)
allstock_info = pd.concat([listed, listedTWO], ignore_index=True)
allstock_info.columns = ["ID","有價證券名稱","市場別","產業別","公開發行/上市(櫃)/發行日"]
allstock_info.set_index('ID', inplace = True)

print(allstock_info)
# print(type(allstock_info['ID'].values[0]))
#==========================================================================
##                                每日更新csv                             ##
#==========================================================================
delay = 0
today=datetime.date.today() - datetime.timedelta(days=delay)
if str(today).split(' ')[0] in HOLIDAY:
    line_notify(f'{today}放假不執行data.py', TOKEN_FOR_UPDATE)
    sys.exit()
line_notify(f'{today}開始執行data.py', TOKEN_FOR_UPDATE)
tomorrow = today + datetime.timedelta(days=1)
time_start = str(today)
time_end = str(tomorrow)
print(f'{bcolors.OK}今日日期:{time_start}{bcolors.RESET}') 
# print(f'{bcolors.OK}明日日期:{time_end}{bcolors.RESET}')
ind_err_num = 0
price_err_num = 0
# ============更新價量&指標============
for i in stock_num :   
    # ============更新價量============
    try:
        stock_data(i, time_start,time_end,delay)
        # print(i + "successful")
    except Exception as e:
        
        try:
            stock_data(i, time_start,time_end,delay)
            # print(i + "successful")
        except Exception as e:
            price_err_num += 1
            print("price fail")
    # ============更新指標============
    try:
        stock_num = indicator(i,allstock_info, stock_num)
        # print(i + "successful")
    except Exception as e:
        
        try:
            stock_num = indicator(i,allstock_info, stock_num)
            # print(i + "successful")
        except Exception as e:
            ind_err_num += 1
            print("indicator fail")
            # print(e)
print(f'{bcolors.FAIL}Price error numbers : {price_err_num} | Indicator error numbers : {ind_err_num}{bcolors.RESET}')
print(f'{bcolors.OK}價量更新完成{bcolors.RESET}')
print(f'{bcolors.OK}指標更新完成{bcolors.RESET}')
line_notify('價量&指標更新完成', TOKEN_FOR_UPDATE)
# ============準備RS rate資料============
timestart = datetime.datetime.strptime(str(datetime.date.today()+datetime.timedelta(days=-delay)) , '%Y-%m-%d' )
rs_df = update_rs_rate(timestart, stock_num)
print(f'{bcolors.OK}rs value參考資料更新完成{bcolors.RESET}')
line_notify('rs value參考資料更新完成', TOKEN_FOR_UPDATE)
rs_df = rs_rate_dataset(rs_df)
print(f'{bcolors.OK}rs rate參考資料更新完成{bcolors.RESET}')
line_notify('rs rate參考資料更新完成', TOKEN_FOR_UPDATE)
# ============將RS rate資料匯入每個csv============
daily_RS_update(timestart, rs_df, stock_num)
print(f'{bcolors.OK}rs rate匯入csv完成{bcolors.RESET}')
line_notify('rs rate匯入csv完成', TOKEN_FOR_UPDATE)
rsmax_err_num = 0
# ============更新RS MAX============
for i in stock_num :      
    try:
        update_RS_MAX(i)
        # print(i + "successful")
    except Exception as e:
        try:
            update_RS_MAX(i)
            # print(i + "successful")
        except Exception as e:
            # print(i + "fail")
            # print(e)
            rsmax_err_num += 1
            pass
print(f'{bcolors.FAIL}RS MAX error numbers : {rsmax_err_num}{bcolors.RESET}')
print(f'{bcolors.OK}rs MAX更新完成{bcolors.RESET}')
line_notify('rs MAX更新完成', TOKEN_FOR_UPDATE)
line_notify(f'{today}data.py執行完成', TOKEN_FOR_UPDATE)