import pandas as pd
import numpy as np
import datetime
import requests
import os
import talib
from io import StringIO
from sklearn.preprocessing import MinMaxScaler
# 設定顏色
class bcolors:
    OK = '\033[92m' #GREEN
    WARNING = '\033[93m' #YELLOW
    FAIL = '\033[91m' #RED
    RESET = '\033[0m' #RESET COLOR

def get_tradingview_format():
    # ============上市股票df============
    url = "https://isin.twse.com.tw/isin/class_main.jsp?owncode=&stockname=&isincode=&market=1&issuetype=1&industry_code=&Page=1&chklike=Y"
    response = requests.get(url, timeout=5)
    print(response.status_code)
    listed = pd.read_html(response.text)[0]
    listed.columns = listed.iloc[0,:]
    listed = listed[["有價證券代號","有價證券名稱","市場別","產業別","公開發行/上市(櫃)/發行日"]]
    listed = listed.iloc[1:]

    # ============上櫃股票df============
    urlTWO = "https://isin.twse.com.tw/isin/class_main.jsp?owncode=&stockname=&isincode=&market=2&issuetype=&industry_code=&Page=1&chklike=Y"
    response = requests.get(urlTWO, timeout=5)
    print(response.status_code)
    listedTWO = pd.read_html(response.text)[0]
    listedTWO.columns = listedTWO.iloc[0,:]
    listedTWO = listedTWO.loc[listedTWO['有價證券別'] == '股票']
    listedTWO = listedTWO[["有價證券代號","有價證券名稱","市場別","產業別","公開發行/上市(櫃)/發行日"]]

    # ============上市股票代號+.TW============
    stock_1 = listed["有價證券代號"]
    stock_num = stock_1.apply(lambda x: str(x) + ".TW")
    stock_num.loc[len(stock_num)+1] = '0050.TW'
    stock_num.loc[len(stock_num)+1] = '^TWII'
    # print(stock_num)

    # ============上櫃股票代號+.TWO============
    stock_2 = listedTWO["有價證券代號"]
    stock_num2 = stock_2.apply(lambda x: str(x) + ".TWO")
    # print(stock_num2)

    # ============concate全部股票代號============
    stock_num = pd.concat([stock_num, stock_num2], ignore_index=True)
    # print(stock_num)
    allstock_info = pd.concat([listed, listedTWO], ignore_index=True)
    allstock_info.columns = ["ID","有價證券名稱","市場別","產業別","公開發行/上市(櫃)/發行日"]
    allstock_info.set_index('ID', inplace = True)
    # print(allstock_info)
    return allstock_info

# =====================data.py=====================
# 刪除資料函式
def delete_data(stock_id, time = 0, dropindex = None):
    address = r'C:\Users\User\Desktop\StockInfoHub\Stock_Data_Collector\history_data\\' + stock_id + ".csv"   
    try:
        if  os.path.isfile(address):
            df_new = pd.read_csv(address,index_col = "Date",parse_dates = ["Date"])
            if time in df_new.index:
                df_new = df_new.drop(time)
                df_new.to_csv(address,encoding='utf-8-sig')
                # print("已刪除指定資料(date)", time)
            elif time == 0:
                df_new = df_new.drop(df_new.index[dropindex])
                df_new.to_csv(address,encoding='utf-8-sig')
                # print("已刪除指定資料(index)",df_new.index[dropindex])
            else:
                pass
                # print('無此日期')
        else:
            print(f'{bcolors.WARNING}{stock_id} 無此檔案{bcolors.RESET}')
        # print(stock_id, 'successful')
    except Exception as err:
        print(f'{bcolors.WARNING}{stock_id}刪除日期失敗({time}, {dropindex}) {err}{bcolors.RESET}')
def download_stock_data(stock_id, s1, s2, max_attempts=5):
    url = f"https://query1.finance.yahoo.com/v7/finance/download/{stock_id}?period1={s1}&period2={s2}&interval=1d&events=history&includeAdjustedClose=true"
    headers = {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36"
    }

    attempts = 0
    while attempts < max_attempts:
        try:
            response = requests.get(url, headers=headers, timeout=10)  # Consider adding a timeout for the request
            response.raise_for_status()  # This will raise an HTTPError if the response was an http error
            return response
        except requests.exceptions.Timeout:
            print(f"Timeout occurred for attempt {attempts + 1}. Retrying...")
            attempts += 1
        except requests.exceptions.HTTPError as err:
            print(f"HTTP Error: {err}")
            return None
        except requests.exceptions.RequestException as e:
            print(f"Error making request: {e}")
            return None

    print("Max attempts reached, failed to make a successful request.")
    return None
##讀寫成csv檔
def stock_data(stock_id,time_start,time_end, delay = 0) :
    days = 24 * 60 * 60    #一天有86400秒 
    initial = datetime.datetime.strptime( '1970-01-01' , '%Y-%m-%d' )
    start = datetime.datetime.strptime( time_start , '%Y-%m-%d' )
    end = datetime.datetime.strptime( time_end, '%Y-%m-%d' )
    period1 = start - initial
    period2 = end - initial
    s1 = period1.days * days
    s2 = period2.days * days
    url = "https://query1.finance.yahoo.com/v7/finance/download/" + stock_id + "?period1=" + str(s1) + "&period2=" + str(s2) + "&interval=1d&events=history&includeAdjustedClose=true"
    headers = {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36"
}
    response = download_stock_data(stock_id, s1, s2)
    df = pd.read_csv(StringIO(response.text),index_col = "Date",parse_dates = ["Date"])
    address = r'C:\Users\User\Desktop\StockInfoHub\Stock_Data_Collector\history_data\\' + stock_id + ".csv"
    ## 刪除舊的當日資料
    deletetime = datetime.datetime.strptime(str(datetime.date.today() - datetime.timedelta(days=delay)) , '%Y-%m-%d' )
    try:
        delete_data(stock_id,deletetime)
    except Exception as e:
        print(e)
    try:
        if  os.path.isfile(address):
            
            df_new = pd.read_csv(address,index_col = "Date",parse_dates = ["Date"])
            if time_start not in df_new.index:
                df_new = pd.concat([df_new,df])
                df_new.to_csv(address,encoding='utf-8')
                # print("已更新到最新資料")
            else:
                print(stock_id, "已是最新價格資料，無需更新")
        else:
            df = df.dropna()
            df.to_csv(address,encoding='utf-8')
            print("此為新資料，已創建csv檔")
    except Exception as err:
        print(f'{bcolors.WARNING}price error,{stock_id}{bcolors.RESET} : {err}')
# 指標更新函式
def indicator(stock_id, allstock_info, stock_num):
    address = r'C:\Users\User\Desktop\StockInfoHub\Stock_Data_Collector\history_data\\' + stock_id + ".csv"   
    comparestock = pd.read_csv(r'C:\Users\User\Desktop\StockInfoHub\Stock_Data_Collector\history_data\\' + '^TWII' + ".csv",index_col = "Date",parse_dates = ["Date"])
    compareindex = comparestock.index.values
    try:
        if  os.path.isfile(address):
            df_new = pd.read_csv(address,index_col = "Date",parse_dates = ["Date"])
            try:
                df_new['ID'] = stock_id.split('.')[0]
            except Exception as e:
                print(e)
                print(stock_id)
            if stock_id not in ['^TWII', '0050.TW']:
                df_new['產業別'] = allstock_info.loc[df_new['ID'].values[0], '產業別']
            df_new['5MA'] = talib.SMA(df_new["Adj Close"], 5)
            df_new['10MA'] = talib.SMA(df_new["Adj Close"], 10)
            df_new['20MA'] = talib.SMA(df_new["Adj Close"], 20)
            df_new['50MA'] = talib.SMA(df_new["Adj Close"], 50)
            df_new['100MA'] = talib.SMA(df_new["Adj Close"], 100)
            df_new['150MA'] = talib.SMA(df_new["Adj Close"], 150)
            df_new['200MA'] = talib.SMA(df_new["Adj Close"], 200)
            df_new['200MA ROCP'] = np.array(map(lambda x : round(x*100,2), talib.ROCP(df_new["200MA"], timeperiod=1)))
            df_new['200MA ROCP 20MA'] = talib.SMA(df_new["200MA ROCP"], 20)
            df_new['200MA ROCP 60MA'] = talib.SMA(df_new["200MA ROCP"], 60)
            df_new['5Max'] = talib.MAX(df_new['High'], 5)
            df_new['250Max'] = talib.MAX(df_new['High'], 250)
            df_new['5Min'] = talib.MIN(df_new['Low'], 5)
            df_new['250Min'] = talib.MIN(df_new['Low'], 250)
            df_new['Volume 50MA'] = talib.SMA(df_new['Volume'], 50)
            # 漲幅
            df_new['ROCP'] = np.array(map(lambda x : round(x*100,2), talib.ROCP(df_new["Adj Close"], timeperiod=1)))
            df_new['OBV'] = talib.OBV(df_new["Adj Close"], df_new["Volume"])
            # ATR
            df_new['ATR250'] = talib.ATR(high=df_new['High'], low=df_new["Low"], close=df_new["Adj Close"], timeperiod=250)
            df_new['ATR50'] = talib.ATR(high=df_new['High'], low=df_new["Low"], close=df_new["Adj Close"], timeperiod=50)
            df_new['ATR20'] = talib.ATR(high=df_new['High'], low=df_new["Low"], close=df_new["Adj Close"], timeperiod=20)
            # std
            df_new['STD7'] = 100*talib.STDDEV(df_new["Adj Close"], timeperiod=7)/talib.EMA(df_new["Adj Close"], 7)
            df_new['STD20'] = 100*talib.STDDEV(df_new["Adj Close"], timeperiod=20)/talib.EMA(df_new["Adj Close"], 20)
            df_new['STD50'] = 100*talib.STDDEV(df_new["Adj Close"], timeperiod=50)/talib.EMA(df_new["Adj Close"], 50)
            stockindex = df_new.index.values
            index = []
            for i in stockindex:
                if i in compareindex:
                    index.append(i)

            beta_df = pd.DataFrame([df_new.loc[index, 'ROCP'].values, comparestock.loc[index, 'ROCP'].values]).transpose()
            beta_df.columns = ['stock', 'TAI']
            beta_df.fillna(1, inplace=True)
            df_new['beta']=0
            beta_list = talib.BETA(beta_df['stock'].values, beta_df['TAI'].values, timeperiod = 50)
            betalistdf = pd.DataFrame(beta_list, index=index, columns=['beta'])
            df_new.loc[index, 'beta'] = betalistdf['beta'].values
            df_new['RS value'] = 0
            df_new.loc[index, 'RS value'] = list(map(lambda stock,compare: stock-compare, df_new.loc[index, 'ROCP'], comparestock.loc[index, 'ROCP']))
            df_new['RS 250MA'] = np.array(talib.SMA(df_new['RS value'], 250))*100
            df_new['RS 50MA'] = np.array(talib.SMA(df_new['RS value'], 50))*100
            df_new['RS 20MA'] = np.array(talib.SMA(df_new['RS value'], 20))*100
            df_new['RS 250EMA'] = np.array(talib.EMA(df_new['RS value'], 250))*100
            df_new['RS 50EMA'] = np.array(talib.EMA(df_new['RS value'], 50))*100
            df_new['RS 20EMA'] = np.array(talib.EMA(df_new['RS value'], 20))*100
            df_new.to_csv(address,encoding='utf-8-sig')
            # print("指標已更新到最新資料")
        else:
            pass
    except Exception as err:
        print(f'{bcolors.WARNING}indicator error,{stock_id}{bcolors.RESET} : {err}')
        try:
            stock_num_list = stock_num.tolist()
            print(f'刪除前數量:{len(stock_num_list)}', end=' | ')
            stock_num_list.remove(stock_id)
            print(f'刪除後數量:{len(stock_num_list)}')
            stock_num = pd.Series(stock_num_list)
        except:
            print('移除指標更新失敗股票，失敗')
    return stock_num
#每日更新RS參考資料
def update_rs_rate(timestart, stock_num):
    rs20rate = []
    first = 1
    rs_df = 0
    for i, id in enumerate(stock_num):
        try:
            stock = pd.read_csv('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/history_data/' + str(id) + '.csv',parse_dates = ["Date"])
            stock.set_index('Date', inplace = True)
            if timestart in stock.index:    
                if first:
                    rs_df = pd.DataFrame(stock.loc[timestart, ['ID', 'RS 250MA', 'RS 50MA', 'RS 20MA', 'RS 250EMA', 'RS 50EMA', 'RS 20EMA']]).transpose()
                    print(rs_df)
                    rs_df.set_index('ID', inplace = True)
                    first = 0
                else:
                    stock = pd.DataFrame(stock.loc[timestart, ['ID', 'RS 250MA', 'RS 50MA', 'RS 20MA', 'RS 250EMA', 'RS 50EMA', 'RS 20EMA']]).transpose()
                    stock.set_index('ID', inplace = True)
                    rs_df = pd.concat([rs_df, stock])
            # print(id, '更新RS參考資料完成')
        except Exception as e:
            print(f'{bcolors.WARNING}{id}{bcolors.RESET} : {e}')
    return rs_df
# 每日做RSrate參考資料
def rs_rate_dataset(rs_df):
    rs_df['RS 250rate'] = 0
    rs_df['RS 50rate'] = 0
    rs_df['RS 20rate'] = 0
    rs_df['RS EMA250rate'] = 0
    rs_df['RS EMA50rate'] = 0
    rs_df['RS EMA20rate'] = 0
    scaler = MinMaxScaler(feature_range=(0, 100))
    try:
        rs_df['RS 250rate'] = rs_df['RS 250MA'].rank(method='dense', ascending=True)
        print(f'{bcolors.OK}RSrate參考資料250排序已完成{bcolors.RESET}')
        rs_df['RS 250rate'] = scaler.fit_transform(np.array(rs_df['RS 250rate']).reshape(-1, 1))
        # print('RSrate參考資料250已完成')
    except Exception as e:
        print(f'{bcolors.WARNING}RSrate參考資料250失敗\n{bcolors.RESET}{e}')
    try:
        rs_df['RS 50rate'] = rs_df['RS 50MA'].rank(method='dense', ascending=True)
        print(f'{bcolors.OK}RSrate參考資料50排序已完成{bcolors.RESET}')
        rs_df['RS 50rate'] = scaler.fit_transform(np.array(rs_df['RS 50rate']).reshape(-1, 1))
        # print('RSrate參考資料50已完成')
    except Exception as e:
        print(f'{bcolors.WARNING}RSrate參考資料50失敗\n{bcolors.RESET}{e}')
    try:
        rs_df['RS 20rate'] = rs_df['RS 20MA'].rank(method='dense', ascending=True)
        print(f'{bcolors.OK}RSrate參考資料20排序已完成{bcolors.RESET}')
        rs_df['RS 20rate'] = scaler.fit_transform(np.array(rs_df['RS 20rate']).reshape(-1, 1))
        # print('RSrate參考資料20已完成')
    except Exception as e:
        print(f'{bcolors.WARNING}RSrate參考資料20失敗\n{bcolors.RESET}{e}')
    try:
        rs_df['RS EMA250rate'] = rs_df['RS 250EMA'].rank(method='dense', ascending=True)
        print(f'{bcolors.OK}EMA RSrate參考資料250排序已完成{bcolors.RESET}')
        rs_df['RS EMA250rate'] = scaler.fit_transform(np.array(rs_df['RS EMA250rate']).reshape(-1, 1))
        # print('RSEMArate參考資料250已完成')
    except Exception as e:
        print(f'{bcolors.WARNING}EMA RSratee參考資料250失敗\n{bcolors.RESET}{e}')
    try:
        rs_df['RS EMA50rate'] = rs_df['RS 50EMA'].rank(method='dense', ascending=True)
        print(f'{bcolors.OK}EMA RSrate參考資料50排序已完成{bcolors.RESET}')
        rs_df['RS EMA50rate'] = scaler.fit_transform(np.array(rs_df['RS EMA50rate']).reshape(-1, 1))
        # print('RSEMArate參考資料50已完成')
    except Exception as e:
        print(f'{bcolors.WARNING}EMA RSrate參考資料50失敗\n{bcolors.RESET}{e}')
    try:
        rs_df['RS EMA20rate'] = rs_df['RS 20EMA'].rank(method='dense', ascending=True)
        print(f'{bcolors.OK}EMA RSrate參考資料20排序已完成{bcolors.RESET}')
        rs_df['RS EMA20rate'] = scaler.fit_transform(np.array(rs_df['RS EMA20rate']).reshape(-1, 1))
        # print('RSEMArate參考資料20已完成')
    except Exception as e:
        print(f'{bcolors.WARNING}EMA RSrate參考資料20失敗\n{bcolors.RESET}{e}')
    return rs_df
# 每日更新RS rate到每個csv
def daily_RS_update(timestart, rs_df, stock_num):
    for i, id in enumerate(stock_num):
        name = id
        address = 'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/history_data/' + str(id) + '.csv'
        try:
            stock = pd.read_csv(address, index_col = "Date",parse_dates = ["Date"])
        except Exception as e:
            print(f'{bcolors.WARNING}{name} {e}{bcolors.RESET}')
        try:
            id = int(id.split('.')[0])
            if timestart in stock.index:
                stock.loc[timestart, ['RS 20rate', 'RS 50rate', 'RS 250rate', 'RS EMA20rate', 'RS EMA50rate', 'RS EMA250rate']] = rs_df.loc[id, ['RS 20rate', 'RS 50rate', 'RS 250rate', 'RS EMA20rate', 'RS EMA50rate', 'RS EMA250rate']]
                stock.to_csv(address,encoding='utf-8-sig')
                # print(name, 'RS rate更新成功')
            else:
                print(name, '沒有RS資料')
        except Exception as e:
            print(f'{bcolors.WARNING}{name}{bcolors.RESET} : {e}')
# 更新RS MAX
def update_RS_MAX(stock_id):
    address = r'C:\Users\User\Desktop\StockInfoHub\Stock_Data_Collector\history_data\\' + stock_id + ".csv"   
    try:
        if  os.path.isfile(address):
            df_new = pd.read_csv(address,index_col = "Date",parse_dates = ["Date"])
            try:
                df_new['ID'] = stock_id.split('.')[0]
            except Exception as e:
                print(e)
                print(f'{bcolors.WARNING}{stock_id} {e}{bcolors.RESET}')
            # 10 max
            df_new['RS 250MA 10MAX'] = talib.MAX(df_new['RS 250rate'], 10)
            df_new['RS 50MA 10MAX'] = talib.MAX(df_new['RS 50rate'], 10)
            df_new['RS 20MA 10MAX'] = talib.MAX(df_new['RS 20rate'], 10)
            df_new['RS 250MA is 10MAX'] = df_new['RS 250rate'].round(1)>=df_new['RS 250MA 10MAX'].round(1)
            df_new['RS 50MA is 10MAX'] = df_new['RS 50rate'].round(1)>=df_new['RS 50MA 10MAX'].round(1)
            df_new['RS 20MA is 10MAX'] = df_new['RS 20rate'].round(1)>=df_new['RS 20MA 10MAX'].round(1)
            df_new['RS 250EMA 10MAX'] = talib.MAX(df_new['RS EMA250rate'], 10)
            df_new['RS 50EMA 10MAX'] = talib.MAX(df_new['RS EMA50rate'], 10)
            df_new['RS 20EMA 10MAX'] = talib.MAX(df_new['RS EMA20rate'], 10)
            df_new['RS 250EMA is 10MAX'] = df_new['RS EMA250rate'].round(1)>=df_new['RS 250EMA 10MAX'].round(1)
            df_new['RS 50EMA is 10MAX'] = df_new['RS EMA50rate'].round(1)>=df_new['RS 50EMA 10MAX'].round(1)
            df_new['RS 20EMA is 10MAX'] = df_new['RS EMA20rate'].round(1)>=df_new['RS 20EMA 10MAX'].round(1)
            # 20 max
            df_new['RS 250MA 20MAX'] = talib.MAX(df_new['RS 250rate'], 20)
            df_new['RS 50MA 20MAX'] = talib.MAX(df_new['RS 50rate'], 20)
            df_new['RS 20MA 20MAX'] = talib.MAX(df_new['RS 20rate'], 20)
            df_new['RS 250MA is 20MAX'] = df_new['RS 250rate'].round(1)>=df_new['RS 250MA 20MAX'].round(1)
            df_new['RS 50MA is 20MAX'] = df_new['RS 50rate'].round(1)>=df_new['RS 50MA 20MAX'].round(1)
            df_new['RS 20MA is 20MAX'] = df_new['RS 20rate'].round(1)>=df_new['RS 20MA 20MAX'].round(1)
            df_new['RS 250EMA 20MAX'] = talib.MAX(df_new['RS EMA250rate'], 20)
            df_new['RS 50EMA 20MAX'] = talib.MAX(df_new['RS EMA50rate'], 20)
            df_new['RS 20EMA 20MAX'] = talib.MAX(df_new['RS EMA20rate'], 20)
            df_new['RS 250EMA is 20MAX'] = df_new['RS EMA250rate'].round(1)>=df_new['RS 250EMA 20MAX'].round(1)
            df_new['RS 50EMA is 20MAX'] = df_new['RS EMA50rate'].round(1)>=df_new['RS 50EMA 20MAX'].round(1)
            df_new['RS 20EMA is 20MAX'] = df_new['RS EMA20rate'].round(1)>=df_new['RS 20EMA 20MAX'].round(1)
            # 50 max
            df_new['RS 250MA 50MAX'] = talib.MAX(df_new['RS 250rate'], 50)
            df_new['RS 50MA 50MAX'] = talib.MAX(df_new['RS 50rate'], 50)
            df_new['RS 20MA 50MAX'] = talib.MAX(df_new['RS 20rate'], 50)
            df_new['RS 250MA is 50MAX'] = df_new['RS 250rate'].round(1)>=df_new['RS 250MA 50MAX'].round(1)
            df_new['RS 50MA is 50MAX'] = df_new['RS 50rate'].round(1)>=df_new['RS 50MA 50MAX'].round(1)
            df_new['RS 20MA is 50MAX'] = df_new['RS 20rate'].round(1)>=df_new['RS 20MA 50MAX'].round(1)
            df_new['RS 250EMA 50MAX'] = talib.MAX(df_new['RS EMA250rate'], 50)
            df_new['RS 50EMA 50MAX'] = talib.MAX(df_new['RS EMA50rate'], 50)
            df_new['RS 20EMA 50MAX'] = talib.MAX(df_new['RS EMA20rate'], 50)
            df_new['RS 250EMA is 50MAX'] = df_new['RS EMA250rate'].round(1)>=df_new['RS 250EMA 50MAX'].round(1)
            df_new['RS 50EMA is 50MAX'] = df_new['RS EMA50rate'].round(1)>=df_new['RS 50EMA 50MAX'].round(1)
            df_new['RS 20EMA is 50MAX'] = df_new['RS EMA20rate'].round(1)>=df_new['RS 20EMA 50MAX'].round(1)
            # 250 max
            df_new['RS 250MA 250MAX'] = talib.MAX(df_new['RS 250rate'], 250)
            df_new['RS 50MA 250MAX'] = talib.MAX(df_new['RS 50rate'], 250)
            df_new['RS 20MA 250MAX'] = talib.MAX(df_new['RS 20rate'], 250)
            df_new['RS 250MA is 250MAX'] = df_new['RS 250rate'].round(1)>=df_new['RS 250MA 250MAX'].round(1)
            df_new['RS 50MA is 250MAX'] = df_new['RS 50rate'].round(1)>=df_new['RS 50MA 250MAX'].round(1)
            df_new['RS 20MA is 250MAX'] = df_new['RS 20rate'].round(1)>=df_new['RS 20MA 250MAX'].round(1)
            df_new['RS 250EMA 250MAX'] = talib.MAX(df_new['RS EMA250rate'], 250)
            df_new['RS 50EMA 250MAX'] = talib.MAX(df_new['RS EMA50rate'], 250)
            df_new['RS 20EMA 250MAX'] = talib.MAX(df_new['RS EMA20rate'], 250)
            df_new['RS 250EMA is 250MAX'] = df_new['RS EMA250rate'].round(1)>=df_new['RS 250EMA 250MAX'].round(1)
            df_new['RS 50EMA is 250MAX'] = df_new['RS EMA50rate'].round(1)>=df_new['RS 50EMA 250MAX'].round(1)
            df_new['RS 20EMA is 250MAX'] = df_new['RS EMA20rate'].round(1)>=df_new['RS 20EMA 250MAX'].round(1)
            df_new.to_csv(address,encoding='utf-8-sig')
            # print("指標已更新到最新資料")
    except Exception as err:
        print(f'{bcolors.WARNING}RS MAX指標更新失敗, {stock_id} {bcolors.RESET} : {err}')

# =====================daily rs indistry.py=====================
# 上市櫃個股代號
def get_allstock_info():
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
    stock_num.loc[len(stock_num)+1] = '0050.TW'
    stock_num.loc[len(stock_num)+1] = '^TWII'
    # print(stock_num)

    # ============上櫃股票代號+.TWO============
    stock_2 = listedTWO["有價證券代號"]
    stock_num2 = stock_2.apply(lambda x: str(x) + ".TWO")
    # print(stock_num2)

    # ============concate全部股票代號============
    stock_num = pd.concat([stock_num, stock_num2], ignore_index=True)
    # print(stock_num)
    allstock_info = pd.concat([listed, listedTWO], ignore_index=True)
    allstock_info.columns = ["ID","有價證券名稱","市場別","產業別","公開發行/上市(櫃)/發行日"]
    allstock_info.set_index('ID', inplace = True)
    return stock_num, allstock_info

# 整合每日個股資料
def concat_stock(day, stock_num):
    allstock = []
    print(f'{bcolors.OK}Stocks Concatenation...{bcolors.RESET}')
    fail_ID = []
    first = 1
    for i, id in enumerate(stock_num):    
        try:
            address = 'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/history_data/' + id + '.csv'
            stockdata = pd.DataFrame(pd.read_csv(address, index_col='Date', parse_dates=['Date']).loc[day]).transpose()
            if first == 1:
                allstock = stockdata
                first = 0
                print(allstock)
            else:
                if len(stockdata.columns) != 93:
                    print(f'{bcolors.WARNING}{id} : columns not match{bcolors.RESET}')
                    print(id)
                    fail_ID.append(id)
                    continue
                allstock = pd.concat([stockdata, allstock], ignore_index=True)
            # print(id)
        except Exception as e:
            print(f'{bcolors.WARNING}{id} : {e}{bcolors.RESET}')
            fail_ID.append(id)
            pass
    try:
        allstock.drop(['RS MA250rate'], axis=1, inplace=True)
    except: 
        pass
    try:
        allstock.drop(['RS MA50rate'], axis=1, inplace=True)
    except:
        pass
    try:
        allstock.drop(['RS MA20rate'], axis=1, inplace=True)
    except:
        pass
    allstock.dropna(inplace=True)
    allstock['ID'] = allstock['ID'].astype(int).astype(str)
    allstock.set_index('ID', inplace = True)
    print(f'{bcolors.OK}DONE{bcolors.RESET}')

    return allstock
# 選股條件
def template(allstock:pd.DataFrame=None, allstock_info:pd.DataFrame=None, yesterday_allstock_info:pd.DataFrame=None):
    print(f'{bcolors.OK}Template Filter...{bcolors.RESET}')
    if yesterday_allstock_info is not None:
        yesterday_allstock_info.set_index('ID', inplace = True)
    for id in allstock.index.values:
        try:
            # stock info
            allstock.loc[id, 'Name'] = allstock_info.loc[str(id), '有價證券名稱']
            allstock.loc[id, 'business volume 50MA(百萬)'] = round(float(allstock.loc[id, 'Volume 50MA'])*float(allstock.loc[id, 'Adj Close'])/1000000, 3)
            allstock.loc[id, 'busness volume(億)'] = (allstock.loc[id, 'Volume'] * allstock.loc[id, 'Adj Close'])/100000000
            allstock.loc[id, 'year high sort'] = (abs((allstock.loc[id, '250Max']-allstock.loc[id, 'Adj Close'])/allstock.loc[id, '250Max'])<0.25)
            allstock.loc[id, 'year low sort'] = ((allstock.loc[id, 'Adj Close']-allstock.loc[id, '250Min'])/allstock.loc[id, '250Min']>0.25)
            # MA strategy
            allstock.loc[id, 'Price>20MA'] = (allstock.loc[id, 'Adj Close']>allstock.loc[id, '20MA'])
            allstock.loc[id, 'Price>50MA'] = (allstock.loc[id, 'Adj Close']>allstock.loc[id, '50MA'])
            allstock.loc[id, 'Price>150MA'] = (allstock.loc[id, 'Adj Close']>allstock.loc[id, '150MA'])
            allstock.loc[id, 'Price>200MA'] = (allstock.loc[id, 'Adj Close']>allstock.loc[id, '200MA'])
            allstock.loc[id, '200MA trending up 60d'] = (allstock.loc[id, '200MA ROCP 60MA'] >0)
            allstock.loc[id, 'RS 250rate>80'] = (allstock.loc[id, 'RS 250rate'] > 80)
            allstock.loc[id, '50MA>150MA'] = (allstock.loc[id, '50MA']>allstock.loc[id, '150MA'])
            allstock.loc[id, '50MA>200MA'] = (allstock.loc[id, '50MA']>allstock.loc[id, '200MA'])
            allstock.loc[id, '150MA>200MA'] = (allstock.loc[id, '150MA']>allstock.loc[id, '200MA'])
            allstock.loc[id, 'price>95%50MA'] = ((allstock.loc[id, 'Adj Close']/allstock.loc[id, '50MA'])>0.95)
            allstock.loc[id, 'price>110%50MA'] = ((allstock.loc[id, 'Adj Close']/allstock.loc[id, '50MA'])>1.1)
            # volume strategy
            allstock.loc[id, 'Volume 50MA>150k'] = (allstock.loc[id, 'Volume 50MA'] > 150*1000)
            allstock.loc[id, 'Volume 50MA>250k'] = (allstock.loc[id, 'Volume 50MA'] > 250*1000)
            allstock.loc[id, 'business volume 50MA(百萬)>200'] = (allstock.loc[id, 'business volume 50MA(百萬)'] > 200)
            # 250 RS strategy
            allstock.loc[id, 'RS 250rate>55'] = (allstock.loc[id, 'RS 250rate'] > 55)
            allstock.loc[id, 'RS 250rate<75'] = (allstock.loc[id, 'RS 250rate'] < 75)
            # 250 ERS strategy
            allstock.loc[id, 'RS EMA250rate>60'] = (allstock.loc[id, 'RS EMA250rate'] > 60)
            allstock.loc[id, 'RS EMA250rate>75'] = (allstock.loc[id, 'RS EMA250rate'] > 75)
            allstock.loc[id, 'RS EMA250rate>80'] = (allstock.loc[id, 'RS EMA250rate'] > 80)
            allstock.loc[id, 'RS EMA250rate>85'] = (allstock.loc[id, 'RS EMA250rate'] > 85)
            allstock.loc[id, 'RS EMA250rate<80'] = (allstock.loc[id, 'RS EMA250rate'] < 80)
            # 20 RS strategy
            allstock.loc[id, 'RS 20rate>85'] = (allstock.loc[id, 'RS 20rate'] > 85)
            # 20 ERS strategy
            allstock.loc[id, 'RS EMA20rate>50'] = (allstock.loc[id, 'RS EMA20rate'] > 50)
            allstock.loc[id, 'RS EMA20rate>80'] = (allstock.loc[id, 'RS EMA20rate'] > 80)
            allstock.loc[id, 'RS EMA20rate<99'] = (allstock.loc[id, 'RS EMA20rate'] < 99)
            # RS diff strategy
            try:
                allstock.loc[id, 'RS EMA20 diff'] = (allstock.loc[id, 'RS EMA20rate'] - yesterday_allstock_info.loc[int(id), 'ES20rate'])
                allstock.loc[id, 'RS EMA20diff < -5'] = (allstock.loc[id, 'RS EMA20 diff'] < -5)
                allstock.loc[id, 'RS EMA20diff < -8'] = (allstock.loc[id, 'RS EMA20 diff'] < -8)
                allstock.loc[id, 'RS EMA20diff < -11'] = (allstock.loc[id, 'RS EMA20 diff'] < -11)
                allstock.loc[id, 'RS EMA20 20MAX diff'] = (allstock.loc[id, 'RS EMA20rate'] - allstock.loc[id, 'RS 20EMA 20MAX'])
                allstock.loc[id, 'RS EMA20 20MAX diff < -5'] = (allstock.loc[id, 'RS EMA20 20MAX diff'] < -5)
                allstock.loc[id, 'RS EMA20 20MAX diff < -10'] = (allstock.loc[id, 'RS EMA20 20MAX diff'] < -10)
                allstock.loc[id, 'RS EMA20 20MAX diff < -20'] = (allstock.loc[id, 'RS EMA20 20MAX diff'] < -20)
            except:
                allstock.loc[id, 'RS EMA20 diff'] = 0
                allstock.loc[id, 'RS EMA20diff < -5'] = False
                allstock.loc[id, 'RS EMA20diff < -8'] = False
                allstock.loc[id, 'RS EMA20diff < -11'] = False
                allstock.loc[id, 'RS EMA20 20MAX diff'] = 0
                allstock.loc[id, 'RS EMA20 20MAX diff < -5'] = False
                allstock.loc[id, 'RS EMA20 20MAX diff < -10'] = False
                allstock.loc[id, 'RS EMA20 20MAX diff < -20'] = False
            # ATR strategy
            allstock.loc[id, 'ATR250/price'] = (allstock.loc[id, 'ATR250'] / allstock.loc[id, 'Adj Close'])
            allstock.loc[id, 'ATR50/price'] = (allstock.loc[id, 'ATR50'] / allstock.loc[id, 'Adj Close'])
            allstock.loc[id, 'ATR20/price'] = (allstock.loc[id, 'ATR20'] / allstock.loc[id, 'Adj Close'])
            allstock.loc[id, 'ATR250/price<0.03'] = (allstock.loc[id, 'ATR250/price'] < 0.03)
            allstock.loc[id, 'ATR250/price<0.5'] = (allstock.loc[id, 'ATR250/price'] < 0.15)
            allstock.loc[id, 'ATR50/price<0.03'] = (allstock.loc[id, 'ATR50/price'] < 0.03)
            allstock.loc[id, 'ATR20/price<0.03'] = (allstock.loc[id, 'ATR20/price'] < 0.03)  
            #All Template
            allstock.loc[id, 'T5'] = all(allstock.loc[id, ['RS 20rate>85', 'RS 250rate>55', 'RS 250rate<75', 'year low sort', 'year high sort', 'Volume 50MA>150k']])
            allstock.loc[id, 'T5-2'] = all(allstock.loc[id, ['RS EMA20rate>80', 'RS EMA250rate>60', 'RS EMA250rate<80', 'year low sort', 'year high sort', 'Volume 50MA>150k']])
            allstock.loc[id, 'T6'] = all(allstock.loc[id, ['RS EMA250rate>80', 'RS EMA20rate>80', 'Volume 50MA>250k']])
            allstock.loc[id, 'T11'] = all(allstock.loc[id, ['RS EMA250rate>75','RS 20EMA is 10MAX','Volume 50MA>250k','price>95%50MA']])
            allstock.loc[id, 'T21'] = all(allstock.loc[id, ['RS EMA20rate>50', 'RS EMA250rate>85', 'RS EMA20 20MAX diff < -5', 'price>110%50MA', 'Volume 50MA>250k']])
            #allstock.loc[id, 'T22'] = all(allstock.loc[id, ['RS EMA20rate>50', 'RS EMA250rate>85', 'RS EMA20 20MAX diff < -10', 'price>110%50MA', 'Volume 50MA>250k']])
            #allstock.loc[id, 'T23'] = all(allstock.loc[id, ['RS EMA20rate>50', 'RS EMA250rate>85', 'RS EMA20 20MAX diff < -20', 'price>110%50MA', 'Volume 50MA>250k']])
            allstock.loc[id, 'TM'] = all(allstock.loc[id, ['Price>150MA', 'Price>200MA', 'year high sort', 'year low sort', '200MA trending up 60d', 'RS 250rate>80', 'Volume 50MA>150k']])
        except Exception as e:
            print(f'{bcolors.WARNING}{id} : {e}{bcolors.RESET}')
            allstock.drop(id, inplace=True)
            # print(id, '失敗')
    print(f'{bcolors.OK}DONE{bcolors.RESET}')
        
    return allstock

# 更新每日產業rs資料
def rs_industry(day):# 
    # print(str(day).split(' ')[0])
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\產業別.xlsx').astype(int).astype(str)
    stock_df = pd.read_excel(f'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選/{str(day).split(" ")[0]}選股.xlsx')
    stock_df = stock_df.sort_values(by='ES250rate', ascending=False).head(340) # 前20%
    stock_id = stock_df['ID'].astype(str)
    number_of_stock = []
    for col in industry_df.columns.values:
        n = len(industry_df.loc[industry_df[col] != '0', col].values)
        number_of_stock.append([col, n])
    number_df = pd.DataFrame(number_of_stock, columns=['industry', 'number'])
    all = []
    for id in stock_id:
        stock_ind = []
        for col in industry_df.columns.values:
            if id in industry_df[col].values:
                stock_ind.append(col)
                all.append([col, id, 1])
    all_df = pd.DataFrame(all, columns=['industry', 'ID', 'count'])
    rs_industry = all_df.groupby(by='industry').sum().sort_values(by='count',ascending=False)
    rs_industry['all number'] = 1
    for industry in rs_industry.index.values:
        rs_industry.loc[industry, 'all number'] = number_df.loc[number_df['industry']==industry, 'number'].values
        rs_industry.loc[industry, 'percentage'] = round(100*rs_industry.loc[industry, 'count']/rs_industry.loc[industry, 'all number'],1)
    rs_industry = rs_industry.transpose()
    daily_rs_percentage = pd.DataFrame(rs_industry.loc['percentage'].values,columns=[str(day)], index=rs_industry.columns.values).transpose()
    history = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/100產業RS排行.xlsx', header=0, index_col=0)
    history.index = history.index.astype(str)
    # if str(day) in history.index.astype(str):
    try:
        history.drop(str(day), inplace=True)
        print(f'{bcolors.WARNING}{str(day)}\nalready update, rewrite today rs industry.{bcolors.RESET}')
    except Exception as e:
        print(e)
        print(f'{bcolors.OK}{str(day)}\nadd today rs industry.{bcolors.RESET}')
    h = pd.concat([history, daily_rs_percentage], axis =0)
    h.sort_index(ascending = False, inplace=True)
    h.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/100產業RS排行.xlsx')
    print(f'{bcolors.OK}rs industry... OK{bcolors.RESET}')
    # else:
    #     print(f'{bcolors.WARNING}{str(day)} already update rs industry{bcolors.RESET}')
# 更新每日產業CF資料
def top_businessvolume_industry(day):
    # print(str(day).split(' ')[0])
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\產業別.xlsx').astype(int).astype(str)
    stock_df = pd.read_excel(f'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選/{str(day).split(" ")[0]}選股.xlsx')
    stock_df = stock_df.sort_values(by='busness volume(億)', ascending=False).head(340) # 前20%
    stock_id = stock_df['ID'].astype(str)
    number_of_stock = []
    for col in industry_df.columns.values:
        n = len(industry_df.loc[industry_df[col] != '0', col].values)
        number_of_stock.append([col, n])
    number_df = pd.DataFrame(number_of_stock, columns=['industry', 'number'])
    all = []
    for id in stock_id:
        stock_ind = []
        for col in industry_df.columns.values:
            if id in industry_df[col].values:
                stock_ind.append(col)
                all.append([col, id, 1])
    all_df = pd.DataFrame(all, columns=['industry', 'ID', 'count'])
    top_volume_industry = all_df.groupby(by='industry').sum().sort_values(by='count',ascending=False)
    top_volume_industry['all number'] = 1
    for industry in top_volume_industry.index.values:
        top_volume_industry.loc[industry, 'all number'] = number_df.loc[number_df['industry']==industry, 'number'].values
        top_volume_industry.loc[industry, 'percentage'] = round(100*top_volume_industry.loc[industry, 'count']/top_volume_industry.loc[industry, 'all number'],1)
    top_volume_industry = top_volume_industry.transpose()
    daily_top_volume_percentage = pd.DataFrame(top_volume_industry.loc['percentage'].values,columns=[str(day)], index=top_volume_industry.columns.values).transpose()
    history = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/100產業成交值排行.xlsx', header=0, index_col=0)
    history.index = history.index.astype(str)
    # if str(day) in history.index.astype(str):
    try:
        history.drop(str(day), inplace=True)
        print(f'{bcolors.WARNING}{str(day)}\nalready update, rewrite today top business volume industry.{bcolors.RESET}')
    except Exception as e:
        print(e)
        print(f'{bcolors.OK}{str(day)}\nadd today top business volume industry.{bcolors.RESET}')
    h = pd.concat([history, daily_top_volume_percentage], axis =0)
    h.sort_index(ascending = False, inplace=True)
    h.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/100產業成交值排行.xlsx')
    print(f'{bcolors.OK}top business volume industry... OK{bcolors.RESET}')
    # else:
    #     print(f'{bcolors.WARNING}{str(day)} already update rs industry{bcolors.RESET}')
# 更新每日產業CF資料(含權重)
def top_businessvolume_industry_with_weight(day):
    # print(str(day).split(' ')[0])
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\產業別.xlsx').astype(int).astype(str)
    stock_df = pd.read_excel(f'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選/{str(day).split(" ")[0]}選股.xlsx')
    stock_df = stock_df.sort_values(by='busness volume(億)', ascending=False).head(340) # 前20%
    stock_id = stock_df['ID'].astype(str)
    number_of_stock = []
    for col in industry_df.columns.values:
        n = len(industry_df.loc[industry_df[col] != '0', col].values)
        number_of_stock.append([col, n])
    number_df = pd.DataFrame(number_of_stock, columns=['industry', 'number'])
    all = []
    for i, id in enumerate(stock_id):
        stock_ind = []
        for col in industry_df.columns.values:
            if id in industry_df[col].values:
                stock_ind.append(col)
                all.append([col, id, 1])
    all_df = pd.DataFrame(all, columns=['industry', 'ID', 'count'])
    top_volume_industry = all_df.groupby(by='industry').sum().sort_values(by='count',ascending=False)
    top_volume_industry['all number'] = 1
    for industry in top_volume_industry.index.values:
        top_volume_industry.loc[industry, 'all number'] = number_df.loc[number_df['industry']==industry, 'number'].values
        top_volume_industry.loc[industry, 'percentage'] = round(100*top_volume_industry.loc[industry, 'count']/top_volume_industry.loc[industry, 'all number'],1)
    top_volume_industry = top_volume_industry.transpose()
    daily_top_volume_percentage = pd.DataFrame(top_volume_industry.loc['percentage'].values,columns=[str(day)], index=top_volume_industry.columns.values).transpose()
    history = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/100產業成交值排行(含權重).xlsx', header=0, index_col=0)
    history.index = history.index.astype(str)
    # if str(day) in history.index.astype(str):
    try:
        history.drop(str(day), inplace=True)
        print(f'{bcolors.WARNING}{str(day)}\nalready update, rewrite today top business volume industry.{bcolors.RESET}')
    except Exception as e:
        print(e)
        print(f'{bcolors.OK}{str(day)}\nadd today top business volume industry.{bcolors.RESET}')
    h = pd.concat([history, daily_top_volume_percentage], axis =0)
    h.sort_index(ascending = False, inplace=True)
    h.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/100產業成交值排行(含權重).xlsx')
    print(f'{bcolors.OK}top business volume industry... OK{bcolors.RESET}')
    # else:
    #     print(f'{bcolors.WARNING}{str(day)} already update rs industry{bcolors.RESET}')

# 更新每日族群rs資料
def rs_group(day):
    # print(str(day).split(' ')[0])
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\族群_複製.xlsx').astype(int).astype(str)
    stock_df = pd.read_excel(f'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選/{str(day).split(" ")[0]}選股.xlsx')
    stock_df = stock_df.sort_values(by='ES250rate', ascending=False).head(340) # 前20%
    stock_id = stock_df['ID'].astype(str)
    number_of_stock = []
    for col in industry_df.columns.values:
        n = len(industry_df.loc[industry_df[col] != '0', col].values)
        number_of_stock.append([col, n])
    number_df = pd.DataFrame(number_of_stock, columns=['industry', 'number'])
    all = []
    for id in stock_id:
        stock_ind = []
        for col in industry_df.columns.values:
            if id in industry_df[col].values:
                stock_ind.append(col)
                all.append([col, id, 1])
    all_df = pd.DataFrame(all, columns=['industry', 'ID', 'count'])
    rs_industry = all_df.groupby(by='industry').sum().sort_values(by='count',ascending=False)
    rs_industry['all number'] = 1
    for industry in rs_industry.index.values:
        
        rs_industry.loc[industry, 'all number'] = number_df.loc[number_df['industry']==industry, 'number'].values
        rs_industry.loc[industry, 'percentage'] = round(100*rs_industry.loc[industry, 'count']/rs_industry.loc[industry, 'all number'],1)
    rs_industry = rs_industry.transpose()
    daily_rs_percentage = pd.DataFrame(rs_industry.loc['percentage'].values,columns=[str(day)], index=rs_industry.columns.values).transpose()
    history = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/族群RS排行.xlsx', header=0, index_col=0)
    history.index = history.index.astype(str)
    # if str(day) in history.index.astype(str):
    try:
        history.drop(str(day), inplace=True)
        print(f'{bcolors.WARNING}{str(day)}\nalready update, rewrite today rs industry.{bcolors.RESET}')
    except Exception as e:
        print(e)
        print(f'{bcolors.OK}{str(day)}\nadd today rs industry.{bcolors.RESET}')
    h = pd.concat([history, daily_rs_percentage], axis =0)
    h.sort_index(ascending = False, inplace=True)
    h.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/族群RS排行.xlsx')
    print(f'{bcolors.OK}rs industry... OK{bcolors.RESET}')
# 更新每日族群CF差異資料
def top_businessvolume_group(day):
    # print(str(day).split(' ')[0])
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\族群_複製.xlsx').astype(int).astype(str)
    stock_df = pd.read_excel(f'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選/{str(day).split(" ")[0]}選股.xlsx')
    stock_df = stock_df.sort_values(by='busness volume(億)', ascending=False).head(340) # 前20%
    stock_id = stock_df['ID'].astype(str)
    number_of_stock = []
    for col in industry_df.columns.values:
        n = len(industry_df.loc[industry_df[col] != '0', col].values)
        number_of_stock.append([col, n])
    number_df = pd.DataFrame(number_of_stock, columns=['industry', 'number'])
    all = []
    for id in stock_id:
        stock_ind = []
        for col in industry_df.columns.values:
            if id in industry_df[col].values:
                stock_ind.append(col)
                all.append([col, id, 1])
    all_df = pd.DataFrame(all, columns=['industry', 'ID', 'count'])
    top_volume_industry = all_df.groupby(by='industry').sum().sort_values(by='count',ascending=False)
    top_volume_industry['all number'] = 1
    for industry in top_volume_industry.index.values:
        top_volume_industry.loc[industry, 'all number'] = number_df.loc[number_df['industry']==industry, 'number'].values
        top_volume_industry.loc[industry, 'percentage'] = round(100*top_volume_industry.loc[industry, 'count']/top_volume_industry.loc[industry, 'all number'],1)
    top_volume_industry = top_volume_industry.transpose()
    daily_top_volume_percentage = pd.DataFrame(top_volume_industry.loc['percentage'].values,columns=[str(day)], index=top_volume_industry.columns.values).transpose()
    history = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/族群成交值排行.xlsx', header=0, index_col=0)
    history.index = history.index.astype(str)
    # if str(day) in history.index.astype(str):
    try:
        history.drop(str(day), inplace=True)
        print(f'{bcolors.WARNING}{str(day)}\nalready update, rewrite today top business volume industry.{bcolors.RESET}')
    except Exception as e:
        print(e)
        print(f'{bcolors.OK}{str(day)}\nadd today top business volume industry.{bcolors.RESET}')
    h = pd.concat([history, daily_top_volume_percentage], axis =0)
    h.sort_index(ascending = False, inplace=True)
    h.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/族群成交值排行.xlsx')
    print(f'{bcolors.OK}top business volume industry... OK{bcolors.RESET}')
    # else:
    #     print(f'{bcolors.WARNING}{str(day)} already update rs industry{bcolors.RESET}')
# 更新每日概念股CF差異資料(含權重)
def top_businessvolume_group_with_weight(day):
    # print(str(day).split(' ')[0])
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\族群_複製.xlsx').astype(int).astype(str)
    stock_df = pd.read_excel(f'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選/{str(day).split(" ")[0]}選股.xlsx')
    stock_df = stock_df.sort_values(by='busness volume(億)', ascending=False).head(340) # 前20%
    stock_id = stock_df['ID'].astype(str)
    number_of_stock = []
    for col in industry_df.columns.values:
        n = len(industry_df.loc[industry_df[col] != '0', col].values)
        number_of_stock.append([col, n])
    number_df = pd.DataFrame(number_of_stock, columns=['industry', 'number'])
    all = []
    for i, id in enumerate(stock_id):
        stock_ind = []
        for col in industry_df.columns.values:
            if id in industry_df[col].values:
                stock_ind.append(col)
                all.append([col, id, 1])
    all_df = pd.DataFrame(all, columns=['industry', 'ID', 'count'])
    top_volume_industry = all_df.groupby(by='industry').sum().sort_values(by='count',ascending=False)
    top_volume_industry['all number'] = 1
    for industry in top_volume_industry.index.values:
        top_volume_industry.loc[industry, 'all number'] = number_df.loc[number_df['industry']==industry, 'number'].values
        top_volume_industry.loc[industry, 'percentage'] = round(100*top_volume_industry.loc[industry, 'count']/top_volume_industry.loc[industry, 'all number'],1)
    top_volume_industry = top_volume_industry.transpose()
    daily_top_volume_percentage = pd.DataFrame(top_volume_industry.loc['percentage'].values,columns=[str(day)], index=top_volume_industry.columns.values).transpose()
    history = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/族群成交值排行(含權重).xlsx', header=0, index_col=0)
    history.index = history.index.astype(str)
    # if str(day) in history.index.astype(str):
    try:
        history.drop(str(day), inplace=True)
        print(f'{bcolors.WARNING}{str(day)}\nalready update, rewrite today top business volume industry.{bcolors.RESET}')
    except Exception as e:
        print(e)
        print(f'{bcolors.OK}{str(day)}\nadd today top business volume industry.{bcolors.RESET}')
    h = pd.concat([history, daily_top_volume_percentage], axis =0)
    h.sort_index(ascending = False, inplace=True)
    h.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/族群成交值排行(含權重).xlsx')
    print(f'{bcolors.OK}top business volume industry... OK{bcolors.RESET}')
    # else:
    #     print(f'{bcolors.WARNING}{str(day)} already update rs industry{bcolors.RESET}')

# 更新每日概念股rs資料
def rs_concept(day):
    # print(str(day).split(' ')[0])
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\概念股_複製.xlsx').astype(int).astype(str)
    stock_df = pd.read_excel(f'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選/{str(day).split(" ")[0]}選股.xlsx')
    stock_df = stock_df.sort_values(by='ES250rate', ascending=False).head(340) # 前20%
    stock_id = stock_df['ID'].astype(str)
    number_of_stock = []
    for col in industry_df.columns.values:
        n = len(industry_df.loc[industry_df[col] != '0', col].values)
        number_of_stock.append([col, n])
    number_df = pd.DataFrame(number_of_stock, columns=['industry', 'number'])
    all = []
    for id in stock_id:
        stock_ind = []
        for col in industry_df.columns.values:
            if id in industry_df[col].values:
                stock_ind.append(col)
                all.append([col, id, 1])
    all_df = pd.DataFrame(all, columns=['industry', 'ID', 'count'])
    rs_industry = all_df.groupby(by='industry').sum().sort_values(by='count',ascending=False)
    rs_industry['all number'] = 1
    for industry in rs_industry.index.values:
        
        rs_industry.loc[industry, 'all number'] = number_df.loc[number_df['industry']==industry, 'number'].values
        rs_industry.loc[industry, 'percentage'] = round(100*rs_industry.loc[industry, 'count']/rs_industry.loc[industry, 'all number'],1)
    rs_industry = rs_industry.transpose()
    daily_rs_percentage = pd.DataFrame(rs_industry.loc['percentage'].values,columns=[str(day)], index=rs_industry.columns.values).transpose()
    history = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/概念股RS排行.xlsx', header=0, index_col=0)
    history.index = history.index.astype(str)
    # if str(day) in history.index.astype(str):
    try:
        history.drop(str(day), inplace=True)
        print(f'{bcolors.WARNING}{str(day)}\nalready update, rewrite today rs industry.{bcolors.RESET}')
    except Exception as e:
        print(e)
        print(f'{bcolors.OK}{str(day)}\nadd today rs industry.{bcolors.RESET}')
    h = pd.concat([history, daily_rs_percentage], axis =0)
    h.sort_index(ascending = False, inplace=True)
    h.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/概念股RS排行.xlsx')
    print(f'{bcolors.OK}rs industry... OK{bcolors.RESET}')
# 更新每日概念股CF差異資料
def top_businessvolume_concept(day):
    # print(str(day).split(' ')[0])
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\概念股_複製.xlsx').astype(int).astype(str)
    stock_df = pd.read_excel(f'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選/{str(day).split(" ")[0]}選股.xlsx')
    stock_df = stock_df.sort_values(by='busness volume(億)', ascending=False).head(340) # 前20%
    stock_id = stock_df['ID'].astype(str)
    number_of_stock = []
    for col in industry_df.columns.values:
        n = len(industry_df.loc[industry_df[col] != '0', col].values)
        number_of_stock.append([col, n])
    number_df = pd.DataFrame(number_of_stock, columns=['industry', 'number'])
    all = []
    for i, id in enumerate(stock_id):
        stock_ind = []
        for col in industry_df.columns.values:
            if id in industry_df[col].values:
                stock_ind.append(col)
                if i<50:
                    all.append([col, id, 3])
                elif i<100:
                    all.append([col, id, 2])
                else:
                    all.append([col, id, 1])
    all_df = pd.DataFrame(all, columns=['industry', 'ID', 'count'])
    top_volume_industry = all_df.groupby(by='industry').sum().sort_values(by='count',ascending=False)
    top_volume_industry['all number'] = 1
    for industry in top_volume_industry.index.values:
        top_volume_industry.loc[industry, 'all number'] = number_df.loc[number_df['industry']==industry, 'number'].values
        top_volume_industry.loc[industry, 'percentage'] = round(100*top_volume_industry.loc[industry, 'count']/top_volume_industry.loc[industry, 'all number'],1)
    top_volume_industry = top_volume_industry.transpose()
    daily_top_volume_percentage = pd.DataFrame(top_volume_industry.loc['percentage'].values,columns=[str(day)], index=top_volume_industry.columns.values).transpose()
    history = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/概念股成交值排行(含權重).xlsx', header=0, index_col=0)
    history.index = history.index.astype(str)
    # if str(day) in history.index.astype(str):
    try:
        history.drop(str(day), inplace=True)
        print(f'{bcolors.WARNING}{str(day)}\nalready update, rewrite today top business volume industry.{bcolors.RESET}')
    except Exception as e:
        print(e)
        print(f'{bcolors.OK}{str(day)}\nadd today top business volume industry.{bcolors.RESET}')
    h = pd.concat([history, daily_top_volume_percentage], axis =0)
    h.sort_index(ascending = False, inplace=True)
    h.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/概念股成交值排行(含權重).xlsx')
    print(f'{bcolors.OK}top business volume industry... OK{bcolors.RESET}')
    # else:
    #     print(f'{bcolors.WARNING}{str(day)} already update rs industry{bcolors.RESET}')
# 更新每日概念股CF差異資料(含權重)
def top_businessvolume_concept_with_weight(day):
    # print(str(day).split(' ')[0])
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\概念股_複製.xlsx').astype(int).astype(str)
    stock_df = pd.read_excel(f'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選/{str(day).split(" ")[0]}選股.xlsx')
    stock_df = stock_df.sort_values(by='busness volume(億)', ascending=False).head(340) # 前20%
    stock_id = stock_df['ID'].astype(str)
    number_of_stock = []
    for col in industry_df.columns.values:
        n = len(industry_df.loc[industry_df[col] != '0', col].values)
        number_of_stock.append([col, n])
    number_df = pd.DataFrame(number_of_stock, columns=['industry', 'number'])
    all = []
    for i, id in enumerate(stock_id):
        stock_ind = []
        for col in industry_df.columns.values:
            if id in industry_df[col].values:
                stock_ind.append(col)
                all.append([col, id, 1])
    all_df = pd.DataFrame(all, columns=['industry', 'ID', 'count'])
    top_volume_industry = all_df.groupby(by='industry').sum().sort_values(by='count',ascending=False)
    top_volume_industry['all number'] = 1
    for industry in top_volume_industry.index.values:
        top_volume_industry.loc[industry, 'all number'] = number_df.loc[number_df['industry']==industry, 'number'].values
        top_volume_industry.loc[industry, 'percentage'] = round(100*top_volume_industry.loc[industry, 'count']/top_volume_industry.loc[industry, 'all number'],1)
    top_volume_industry = top_volume_industry.transpose()
    daily_top_volume_percentage = pd.DataFrame(top_volume_industry.loc['percentage'].values,columns=[str(day)], index=top_volume_industry.columns.values).transpose()
    history = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/概念股成交值排行.xlsx', header=0, index_col=0)
    history.index = history.index.astype(str)
    # if str(day) in history.index.astype(str):
    try:
        history.drop(str(day), inplace=True)
        print(f'{bcolors.WARNING}{str(day)}\nalready update, rewrite today top business volume industry.{bcolors.RESET}')
    except Exception as e:
        print(e)
        print(f'{bcolors.OK}{str(day)}\nadd today top business volume industry.{bcolors.RESET}')
    h = pd.concat([history, daily_top_volume_percentage], axis =0)
    h.sort_index(ascending = False, inplace=True)
    h.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/概念股成交值排行.xlsx')
    print(f'{bcolors.OK}top business volume industry... OK{bcolors.RESET}')
    # else:
    #     print(f'{bcolors.WARNING}{str(day)} already update rs industry{bcolors.RESET}')
# 更新每日策略選股數量
def strategy_stock_number(day):
    folder_name = 'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選'
    file_names = day+'選股'
    print(file_names)
    industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\產業別.xlsx').astype(int).astype(str)
    group_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\族群_複製.xlsx').astype(int).astype(str)
    concept_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\概念股_複製.xlsx').astype(int).astype(str)
    TAIEX_df = pd.read_csv(r'C:\Users\User\Desktop\StockInfoHub\Stock_Data_Collector\history_data\^TWII.csv')

    daily_template_industry_file = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/每日策略選股產業數量.xlsx')
    daily_template_group_file = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/每日策略選股族群數量.xlsx')
    daily_template_concept_file = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/每日策略選股概念股數量.xlsx')
    daily_template_df_file = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/每日策略選股數量.xlsx')


    all_industry_names = industry_df.columns.tolist()
    all_industry_names = np.concatenate([all_industry_names, group_df.columns.tolist()])
    all_industry_names = np.concatenate([all_industry_names, concept_df.columns.tolist()])
    all_industry_names = np.unique(all_industry_names).tolist()
    templates = ['T5', 'T5-2', 'T6', 'T11', 'T21', 'CF']

    daily_template_industry = pd.DataFrame(np.zeros((len(file_names), len(templates)+1)), columns=['Date'] + templates)
    daily_template_group = pd.DataFrame(np.zeros((len(file_names), len(templates)+1)), columns=['Date'] + templates)
    daily_template_concept = pd.DataFrame(np.zeros((len(file_names), len(templates)+1)), columns=['Date'] + templates)
    daily_template_df = pd.DataFrame(np.zeros((len(file_names), len(templates)+3)), columns=['Date'] + templates + ['TAIEX', 'TAIEX change %'])
    for i, file in enumerate([file_names]):
        file_path = folder_name + '/' + file + '.xlsx'
        df = pd.read_excel(file_path)
        ids = df['ID'].astype(str)
        if 'ID' in df.columns:
            df['ID'] = df['ID'].astype(str)
            df.set_index('ID', inplace = True)
            print(f'{bcolors.OK} {file} {bcolors.RESET} is loaded.')
        else:
            print(f'{bcolors.FAIL} {file} {bcolors.RESET} is not loaded.')
            continue
        try:
            a = allstock_info
        except NameError:
            allstock_info = get_tradingview_format()
        else:
            pass
        #看個股產業
        alist = []
        industry_category_df = []
        show = True
        df = df.sort_values(by='busness volume(億)', ascending=False)
        df['CF'] = [True if index_value < 50 else False for index_value, x in enumerate(df.index.values)]
        date = file.replace('選股', '')
        daily_template_industry.iloc[i, 0] = date
        daily_template_group.iloc[i, 0] = date
        daily_template_concept.iloc[i, 0] = date
        daily_template_df.iloc[i, 0] = date
        daily_template_df.loc[i, 'TAIEX'] = TAIEX_df[TAIEX_df['Date'] == date]['Close'].values[0]
        daily_template_df.loc[i, 'TAIEX change %'] = TAIEX_df[TAIEX_df['Date'] == date]['ROCP'].values[0]
        for template_index, template in enumerate(templates):
            to_watch = df[df[template] == True].index.values
            all_stock_industry = []
            all_stock_group = []
            all_stock_concept = []
            daily_template_df.loc[i, f'{template}'] = len(to_watch)
            for id in to_watch:
                stock_industry = []
                stock_group = []
                stock_concept = []
                for col in industry_df.columns.values:
                    if id in industry_df[col].values:
                        stock_industry.append(col)
                for col in group_df.columns.values:
                    if id in group_df[col].values:
                        stock_group.append(col)
                for col in concept_df.columns.values:
                    if id in concept_df[col].values:
                        stock_concept.append(col)
                stock_industry = np.unique(stock_industry).tolist()
                stock_group = np.unique(stock_group).tolist()
                stock_concept = np.unique(stock_concept).tolist()
                all_stock_industry += stock_industry
                all_stock_group += stock_group
                all_stock_concept += stock_concept
            def count_industry(all_stock_industry):
                count_of_industry = pd.DataFrame(all_stock_industry, columns=['industry'])
                count_of_industry['count'] = 1
                count_of_industry = count_of_industry.groupby('industry').sum()
                count_of_industry = count_of_industry.sort_values(by='count', ascending=False)
                count_of_industry = count_of_industry.iloc[:3]
                text = ''
                for index in count_of_industry.index.values:
                    text += f"{str(index)}({count_of_industry.loc[index, 'count'].astype(str)}) / "
                
                return text
            daily_template_industry.iloc[i, template_index+1] = count_industry(all_stock_industry)
            daily_template_group.iloc[i, template_index+1] = count_industry(all_stock_group)
            daily_template_concept.iloc[i, template_index+1] = count_industry(all_stock_concept)

    daily_template_concept_file = pd.concat([daily_template_concept_file, daily_template_concept], ignore_index=True)
    daily_template_group_file = pd.concat([daily_template_group_file, daily_template_group], ignore_index=True)
    daily_template_industry_file = pd.concat([daily_template_industry_file, daily_template_industry], ignore_index=True)
    daily_template_df_file = pd.concat([daily_template_df_file, daily_template_df], ignore_index=True)    

    daily_template_industry_file.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/每日策略選股產業數量.xlsx', index=False)
    daily_template_group_file.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/每日策略選股族群數量.xlsx', index=False)
    daily_template_concept_file.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/每日策略選股概念股數量.xlsx', index=False)
    daily_template_df_file.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/每日策略選股數量.xlsx', index=False)


# 更新每日策略選到股票差異
def diff_strategy_stock(industry_df, all_add_ID, all_drop_ID, number_of_stock_in_industry:pd.DataFrame):
    add_industry_df = print_industry_category_df(industry_df,all_add_ID,[],number_of_stock_in_industry, t='新增')
    drop_industry_df = print_industry_category_df(industry_df,all_drop_ID,[],number_of_stock_in_industry, t='刪除')
    all_industry_df = pd.concat([add_industry_df, drop_industry_df], axis=0)
    all_industry_df.fillna(0, inplace=True)
    all_industry_df.loc['合計新增數量'] = all_industry_df.loc['該類股新增數量'] - all_industry_df.loc['該類股刪除數量']
    all_industry_df.loc['合計新增股票%數'] = all_industry_df.loc['新增股票%數'] - all_industry_df.loc['刪除股票%數']
    all_industry_df = all_industry_df.transpose()
    all_industry_df.sort_values(by='合計新增股票%數', ascending=False, inplace=True)
    # all_industry_df['新增股票數量排名'] = all_industry_df['該類股新增數量'].rank(ascending=False)
    add_industry_df = all_industry_df[['該類股新增數量', '該類股刪除數量', '合計新增數量', '合計新增股票%數']].iloc[:30]
    add_industry_df['合計新增數量排名'] = add_industry_df['合計新增數量'].rank(ascending=False)
    add_industry_df['該類股總數量'] = number_of_stock_in_industry.loc[add_industry_df.index.values]
    add_industry_df = add_industry_df.loc[add_industry_df['合計新增數量']>0]
    add_industry_df.sort_values(by='合計新增股票%數', ascending=False, inplace=True)
    add_industry_df = add_industry_df.iloc[:, [5,0,1,2,3,4]].transpose()
    add_industry_df.iloc[[0,1,2,3,5]] = add_industry_df.iloc[[0,1,2,3,5]].astype(int)
    drop_industry_df = all_industry_df[['該類股新增數量', '該類股刪除數量', '合計新增數量', '合計新增股票%數']].iloc[-30:]
    drop_industry_df['合計刪除數量排名'] = drop_industry_df['合計新增數量'].rank(ascending=True)
    drop_industry_df['該類股總數量'] = number_of_stock_in_industry.loc[drop_industry_df.index.values]
    drop_industry_df = drop_industry_df.loc[drop_industry_df['合計新增數量']<0]
    drop_industry_df.sort_values(by='合計新增股票%數', ascending=True, inplace=True)
    drop_industry_df = drop_industry_df.iloc[:, [5,0,1,2,3,4]].transpose()
    drop_industry_df.iloc[[0,1,2,3,5]] = drop_industry_df.iloc[[0,1,2,3,5]].astype(int)
    # ============email新增股票產業的表格文字============
    html_addstock_table = add_industry_df.to_html(index=True, header=True, justify='center')
    html_dropstock_table = drop_industry_df.to_html(index=True, header=True, justify='center')
    return html_addstock_table, html_dropstock_table
# 更新每日產業rs差異資料
def print_industry_category_df(industry_df, ID_list, industry_category_df, number_of_stock_in_industry:pd.DataFrame, print_all = False, t = '新增'):
            alist = []
            industry_category_df = []
            for id in ID_list.astype(str):
                stock_ind = []
                for col in industry_df.columns.values:
                    if id in industry_df[col].values:
                        stock_ind.append(col)
                alist.append((id,stock_ind))

            for ind in alist:
                for i in ind[1]:
                    industry_category_df.append([ind[0], i])
            industry_category_df = pd.DataFrame(industry_category_df,columns=['ID', 'category'])
            industry_category_df['count'] = 1
            industry_category_df = industry_category_df.groupby('category').sum()
            industry_category_df = industry_category_df.sort_values(by='count', ascending=False)
            number_of_stock_in_industry = number_of_stock_in_industry.loc[list(map(lambda x:x in industry_category_df.index.values, number_of_stock_in_industry.index.values))]
            industry_category_df = pd.concat([industry_category_df, number_of_stock_in_industry], axis=1)
            industry_category_df['percentage'] = 100*industry_category_df['count']/industry_category_df['number']
            industry_category_df['percentage'] = industry_category_df['percentage'].round(2)
            industry_category_df['count rank'] = industry_category_df['count'].rank(ascending=False)
            industry_category_df[['count rank', 'count', 'number']] = industry_category_df[['count rank', 'count', 'number']].astype(int)
            industry_category_df = industry_category_df.sort_values(by='percentage', ascending=False).transpose()
            # industry_category_df = industry_category_df.loc[:, industry_category_df.loc['count']>1]
            industry_category_df.index = [f'該類股{t}數量', '該類股總數量', f'{t}股票%數', f'{t}股票數量排名']
            seperate_n = 13
            # for i in range(int(len(industry_category_df.columns.values)/seperate_n)+1):
            #     display.display(industry_category_df.iloc[[2,3,0,1], i*seperate_n:(i+1)*seperate_n])
            # if print_all:
            #     for ind in alist:
            #         print(ind[0], ':', end=' ')
            #         for i in ind[1]:
            #             print(i, end=' ')
            #         print('')
            return industry_category_df
