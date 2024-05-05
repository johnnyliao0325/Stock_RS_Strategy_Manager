import datetime
import requests
import pandas as pd
import numpy as np
import warnings
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
import openpyxl
import traceback
import sys
from tqdm import tqdm
sys.path.append('C:/Users/User/Desktop/StockInfoHub')
from Shared_Modules.shared_functions import *
from Shared_Modules.shared_variables import *
from stock_data_collector_package.stock_data_collector_utils import *
warnings.filterwarnings('ignore')
class bcolors:
    OK = '\033[92m' #GREEN
    WARNING = '\033[93m' #YELLOW
    FAIL = '\033[91m' #RED
    RESET = '\033[0m' #RESET COLOR
#=============是否寄email============
sentemailornot = True
istest = False
update_industry_excel = True
excel_description = ''
# ============上市股票df============
stock_num, allstock_info = get_allstock_info()
# print(type(allstock_info['ID'].values[0]))
print(allstock_info)
#============ 把要看日期的個股資料合併(一天)============
# for i in tqdm(range(2000,0,-1), postfix='*******', ncols=150):
for i in [0]:
# ============改日期============
    # 今天日期往前推幾天(今日 - n_day_ago)
    n_day_ago = 0
    try:
        day = datetime.datetime.strptime(str(datetime.date.today() - datetime.timedelta(days=n_day_ago)) , '%Y-%m-%d' )
        if str(day).split(' ')[0] in HOLIDAY:
            line_notify(f'{day}放假不執行daily_rs_industry.py', TOKEN_FOR_UPDATE)
            continue
            # sys.exit()
        ## 假日不執行
        if day.weekday() == 5 or day.weekday() == 6:
            line_notify(f'{day}放假不執行daily_rs_industry.py', TOKEN_FOR_UPDATE)
            continue
        print('day: ', day)
        line_notify(f'{day}開始執行daily_rs_industry.py', TOKEN_FOR_UPDATE)
        yesterday_allstock_info = None
        # ============讀取昨日選股============
        for y_i in range(10):
            try:
                d = datetime.datetime.strptime(str(datetime.date.today() - datetime.timedelta(days=y_i+1+n_day_ago)), '%Y-%m-%d')
                d = str(d).split(' ')[0]
                yesterday_allstock_info = pd.read_excel(fr'C:\Users\User\Desktop\StockInfoHub\Stock_Data_Collector\全個股條件篩選\{d}選股.xlsx')
                print(f'{bcolors.OK}yesterday_allstock_info {d} OK{bcolors.RESET}')
                break
            except:
                continue

        ## ============第一個日期選出符合條件的股票並匯出excel============
        allstock = concat_stock(day, stock_num[2:])
        line_notify('concat_stock OK', TOKEN_FOR_UPDATE)
        # fail_ID = ['1101.TW', '1102.TW', '1103.TW', '1104.TW', '1108.TW', '1109.TW', '1201.TW', '1203.TW', '1213.TW', '1229.TW', '1234.TW', '1313.TW', '1435.TW', '1733.TW', '2434.TW', '2440.TW', '2528.TW', '4569.TW', '6225.TW', '6655.TW', '2740.TWO', '2752.TWO', '3085.TWO', '3226.TWO', '3332.TWO', '3426.TWO', '3523.TWO', '4192.TWO', '4556.TWO', '4568.TWO', '4712.TWO', '4806.TWO', '5016.TWO', '5455.TWO', '5520.TWO', '5601.TWO', '6103.TWO', '6219.TWO', '6236.TWO', '6629.TWO', '8077.TWO', '8342.TWO', '8423.TWO', '8472.TWO', '8917.TWO', '8921.TWO', '9960.TWO']
        # fail_ID = list(map(lambda x:x.split('.')[0], fail_ID))
        # for failid in fail_ID:
        #     try:
        #         allstock.drop(failid, inplace=True)
        #     except:
                # pass

        # ============加入是否符合策略============
        allstock = template(allstock,allstock_info,yesterday_allstock_info)
        line_notify('template OK', TOKEN_FOR_UPDATE)
        # ID = allstock.index.values
        allstock.dropna(inplace=True)
        T5_ID = allstock.loc[allstock['T5']].index.values.astype(str)
        T5_2_ID = allstock.loc[allstock['T5-2']].index.values.astype(str)
        T6_ID = allstock.loc[allstock['T6'], ['產業別', 'RS EMA250rate']].sort_values(by='RS EMA250rate', ascending=False).iloc[:50].index.values.astype(str)
        T11_ID = allstock.loc[allstock['T11']].index.values.astype(str)
        T21_ID = allstock.loc[allstock['T21']].index.values.astype(str)
        TM_ID = allstock.loc[allstock['TM']].index.values.astype(str)
        
        T5_stock_num = len(T5_ID)
        T5_2_stock_num = len(T5_2_ID)
        T6_stock_num = len(T6_ID)
        T11_stock_num = len(T11_ID)
        T21_stock_num = len(T21_ID)
        TM_stock_num = len(TM_ID)
        all_template_ID = list(set(np.concatenate([T5_ID,T5_2_ID,T6_ID,T11_ID,TM_ID,T21_ID])))
        all_template_ID_exTM = list(set(np.concatenate([T5_ID,T5_2_ID,T6_ID,T11_ID,T21_ID])))
        print('T5 股票數量 : ', T5_stock_num, end=' | ')
        print('T5-2 股票數量 : ', T5_2_stock_num, end=' | ')
        print('T6 股票數量 : ', T6_stock_num, end=' | ')
        print('T11 股票數量 : ', T11_stock_num, end=' | ')
        print('T21 股票數量 : ', T21_stock_num, end=' | ')
        print('TM 股票數量 : ', TM_stock_num)
        print('除了TM以外的股票數量 : ', len(all_template_ID))
        print('全部股票數量 : ', len(all_template_ID_exTM))
        
        # ============防止用濾網篩選前50個的template混亂============
        allstock[['T5','T5-2','T6','T11','TM','T21']] = False
        allstock.loc[T5_ID,'T5'] = True
        allstock.loc[T5_2_ID,'T5-2'] = True
        allstock.loc[T6_ID,'T6'] = True
        allstock.loc[T11_ID,'T11'] = True
        allstock.loc[T21_ID,'T21'] = True
        allstock.loc[TM_ID,'TM'] = True
        allstock[['RS 250rate', 'RS 50rate', 'RS 20rate', 'RS EMA250rate', 'RS EMA50rate', 'RS EMA20rate']] = allstock[['RS 250rate', 'RS 50rate', 'RS 20rate', 'RS EMA250rate', 'RS EMA50rate', 'RS EMA20rate']].astype(float).round(1)
        allstock['Volume 50MA'] = (allstock['Volume 50MA']/1000).astype(int)
        # for i,j in enumerate(allstock['TM']):
        #     if j != j:
        #         print(j, i, ID[i])
        #         allstock.drop(ID[i], axis=0, inplace=True)
        print(f'{bcolors.OK}All Template Stocks... {bcolors.RESET}', end='')

        #============今日符合策略全部股票============
        get_column_name = ['Name', 'Adj Close', 'busness volume(億)', 'Volume 50MA', 'business volume 50MA(百萬)', '產業別'
                                            , 'RS EMA250rate', 'RS EMA50rate', 'RS EMA20rate', 'RS 250rate', 'RS 50rate', 'RS 20rate'
                                            , '5MA', '10MA', '20MA', '50MA', '100MA', '150MA', '200MA', 'ROCP'
                                            , 'ATR250/price', 'ATR50/price', 'ATR20/price', 'STD7', 'STD20', 'STD50', 'RS EMA20 diff', 'RS EMA20 20MAX diff'
                                            ,'RS 250EMA is 10MAX','RS 50EMA is 10MAX','RS 20EMA is 10MAX', 'T5', 'T5-2', 'T6', 'T11', 'TM', 'T21'
                                            , 'Price>150MA', 'Price>200MA', '50MA>150MA', '50MA>200MA', '150MA>200MA','year high sort', 'year low sort'
                                            , '200MA trending up 60d', 'Volume 50MA>150k', 'Volume 50MA>250k']
        
        rename_column_name = ['Name', 'Adj Close', 'busness volume(億)', 'Volume 50MA(張)', 'business volume 50MA(百萬)', '產業別'
                , 'ES250rate', 'ES50rate', 'ES20rate', 'S250rate', 'S50rate', 'S20rate'
                , '5MA', '10MA', '20MA', '50MA', '100MA', '150MA', '200MA', 'ROCP'
                , 'ATR250/price', 'ATR50/price', 'ATR20/price', 'STD7', 'STD20', 'STD50', 'RS EMA20 diff', 'RS EMA20 20MAX diff'
                ,'ES250 is 10D MAX','ES50 is 10D MAX','ES20 is 10D MAX', 'T5', 'T5-2', 'T6', 'T11', 'TM', 'T21'
                , 'Price>150MA', 'Price>200MA', '50MA>150MA', '50MA>200MA', '150MA>200MA','year high sort', 'year low sort'
                , '200MA trending up 60d', 'Volume 50MA>150k', 'Volume 50MA>250k']
        
        apexstock = allstock.loc[all_template_ID, get_column_name]
        apexstock.columns = rename_column_name
        apexstock = apexstock.dropna()
        print(f'{bcolors.OK}OK{bcolors.RESET}')
        line_notify('all template stocks OK', TOKEN_FOR_UPDATE)
        print(f'All Template Stocks shape:{apexstock.shape}')
        
        print(f'{bcolors.OK}RS EMA250rate>80 Stocks... {bcolors.RESET}', end='')

        #============今日策略250ERS>80股票============
        apexstock2 = allstock.loc[allstock['RS EMA250rate>80'], get_column_name]
        apexstock2.columns = rename_column_name
        apexstock2 = apexstock2.dropna()
        print(f'{bcolors.OK}OK{bcolors.RESET}')
        line_notify('RS EMA 250rate>80 stocks OK', TOKEN_FOR_UPDATE)
        print(f'RS EMA250rate>80 Stocks shape:{apexstock2.shape}')

        print(f'{bcolors.OK}All Stocks... {bcolors.RESET}', end='')

        #============今日全部股票============
        sentstock = allstock.copy()
        allstock = allstock.loc[:, get_column_name]
        allstock.columns = rename_column_name
        allstock = allstock.dropna()
        print(f'{bcolors.OK}OK{bcolors.RESET}')
        line_notify('all stocks OK', TOKEN_FOR_UPDATE)
        print(f'All Stocks Shape:{allstock.shape}{bcolors.RESET}')
        
        # print(len(apexstock.index))

        #============匯出全個股條件篩選excel============
        sent_file_path = 'C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/全個股條件篩選/' + str(day).split(' ')[0] + f'選股{excel_description}' + '.xlsx'
        allstock.to_excel(sent_file_path, encoding='utf-8-sig')

        #============匯出符合策略股票excel============
        apexstock.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/250RS選股/' + str(day).split(' ')[0] + f'250RS選股{excel_description}' + '.xlsx', encoding='utf-8-sig')

        #============匯出250ERS>80股票excel============
        apexstock2.to_excel('C:/Users/User/Desktop/StockInfoHub/Stock_Data_Collector/250RS選股/' + str(day).split(' ')[0] + f'EMA250RS選股{excel_description}' + '.xlsx', encoding='utf-8-sig')
        if update_industry_excel:
            try:
                rs_industry(day)
            except:
                line_notify('⚠️rs_industry FAIL', TOKEN_FOR_UPDATE)
                sys.exit()
            line_notify('rs_industry OK', TOKEN_FOR_UPDATE)
            try:
                top_businessvolume_industry(day)
            except:
                line_notify('⚠️top_volume_industry FAIL', TOKEN_FOR_UPDATE)
                sys.exit()
            line_notify('top_volume_industry OK', TOKEN_FOR_UPDATE)
            try:
                top_businessvolume_industry_with_weight(day)
            except:
                line_notify('⚠️top_volume_industry_with_weight FAIL', TOKEN_FOR_UPDATE)
                sys.exit()
            line_notify('top_volume_industry_with_weight OK', TOKEN_FOR_UPDATE)
            try:
                rs_group(day)
            except:
                line_notify('⚠️rs_group FAIL', TOKEN_FOR_UPDATE)
                sys.exit()
            line_notify('rs_group OK', TOKEN_FOR_UPDATE)
            try:
                top_businessvolume_group(day)
            except:
                line_notify('⚠️top_volume_group FAIL', TOKEN_FOR_UPDATE)
                sys.exit()
            line_notify('top_volume_group OK', TOKEN_FOR_UPDATE)
            try:
                top_businessvolume_group_with_weight(day)
            except:
                line_notify('⚠️top_volume_group_with_weight FAIL', TOKEN_FOR_UPDATE)
                sys.exit()
            line_notify('top_volume_group_with_weight OK', TOKEN_FOR_UPDATE)
            try:
                rs_concept(day)
            except:
                line_notify('⚠️rs_concept FAIL', TOKEN_FOR_UPDATE)
                sys.exit()
            line_notify('rs_concept OK', TOKEN_FOR_UPDATE)
            try:
                top_businessvolume_concept(day)
            except:
                line_notify('⚠️top_volume_concept FAIL', TOKEN_FOR_UPDATE)
                sys.exit()
            line_notify('top_volume_concept OK', TOKEN_FOR_UPDATE)
            try:
                top_businessvolume_concept_with_weight(day)
            except:
                line_notify('⚠️top_volume_concept_with_weight FAIL', TOKEN_FOR_UPDATE)
                sys.exit()
            line_notify('top_volume_concept_with_weight OK', TOKEN_FOR_UPDATE)
            try:
                strategy_stock_number(str(day).split(' ')[0])
            except:
                line_notify('⚠️strategy_stock_number FAIL', TOKEN_FOR_UPDATE)
                sys.exit()
            print(str(day))
            
            ## clean terminal
            
        ###############################################################################################################################################################################
        ##                                                                                寄email                                                                                    ##
        ############################################################################################################################################################################### 
        
        if sentemailornot:
            ers_higher80 = sentstock.loc[sentstock['RS EMA250rate>80']]
            ers_higher80_length = len(ers_higher80.index)
            all_length = len(sentstock.index)
            numb_Price_higher_20MA = len(ers_higher80.loc[ers_higher80['Price>20MA'], 'Price>20MA'].index.values)
            numb_Price_higher_50MA = len(ers_higher80.loc[ers_higher80['Price>50MA'], 'Price>50MA'].index.values)
            numb_Price_higher_150MA = len(ers_higher80.loc[ers_higher80['Price>150MA'], 'Price>150MA'].index.values)
            numb_Price_higher_200MA = len(ers_higher80.loc[ers_higher80['Price>200MA'], 'Price>200MA'].index.values)
            all_numb_Price_higher_20MA = len(sentstock.loc[sentstock['Price>20MA'], 'Price>20MA'].index.values)
            all_numb_Price_higher_50MA = len(sentstock.loc[sentstock['Price>50MA'], 'Price>50MA'].index.values)
            all_numb_Price_higher_150MA = len(sentstock.loc[sentstock['Price>150MA'], 'Price>150MA'].index.values)
            all_numb_Price_higher_200MA = len(sentstock.loc[sentstock['Price>200MA'], 'Price>200MA'].index.values)
            print(f'{bcolors.OK}寄送email...{bcolors.RESET}')
            

            ## 比較今天的選股跟昨天的選股
            
            # ============讀今日和昨日的股票Excel============
            path = r'C:\Users\User\Desktop\StockInfoHub\Stock_Data_Collector\全個股條件篩選\\'
            new_date = str(day).split(' ')[0]
            delay_day = 0
            while True:
                delay_day += 1
                old_date = str(datetime.datetime.strptime(str(datetime.date.today() - datetime.timedelta(days=n_day_ago+delay_day)) , '%Y-%m-%d' )).split(' ')[0]
                print(f'New date:{new_date} | Old date:{old_date}')
                try:
                    df_old = pd.read_excel(f'{path}{old_date}選股.xlsx')
                    df_new = pd.read_excel(f'{path}{new_date}選股.xlsx')
                    df_old.set_index('ID', inplace=True)
                    df_new.set_index('ID', inplace=True)
                    break
                except:
                    continue            
            print(f'New date:{new_date} | Old date:{old_date}')

            # ============今日和昨日股票的ID============
            df_same_ID = df_new.loc[df_new.index.isin(df_old.index.values)].index.values
            df_old = df_old.loc[df_same_ID]
            df_new = df_new.loc[df_same_ID]
            old_T5_ID = df_old.loc[df_old['T5']].index.values
            old_T5_2_ID = df_old.loc[df_old['T5-2']].index.values
            old_T6_ID = df_old.loc[df_old['T6']].index.values
            old_T11_ID = df_old.loc[df_old['T11']].index.values
            new_T5_ID = df_new.loc[df_new['T5']].index.values
            new_T5_2_ID = df_new.loc[df_new['T5-2']].index.values
            new_T6_ID = df_new.loc[df_new['T6']].index.values
            new_T11_ID = df_new.loc[df_new['T11']].index.values

            # ============新增和刪除股票的ID============
            T5_add_ID = np.sort([i for i in new_T5_ID if i not in old_T5_ID])
            T5_drop_ID = np.sort([i for i in old_T5_ID if i not in new_T5_ID])
            T6_add_ID = np.sort([i for i in new_T6_ID if i not in old_T6_ID])
            T6_drop_ID = np.sort([i for i in old_T6_ID if i not in new_T6_ID])
            T11_add_ID = np.sort([i for i in new_T11_ID if i not in old_T11_ID])
            T11_drop_ID = np.sort([i for i in old_T11_ID if i not in new_T11_ID])
            T5_2_add_ID = np.sort([i for i in new_T5_2_ID if i not in old_T5_2_ID])
            T5_2_drop_ID = np.sort([i for i in old_T5_2_ID if i not in new_T5_2_ID])
            T5_drop_goodID = np.sort([i for i in T5_drop_ID if all([df_new.loc[i,'S250rate']>=75, df_new.loc[i, 'S20rate']>85])])
            T5_2_drop_goodID = np.sort([i for i in T5_2_drop_ID if all([df_new.loc[i,'ES250rate']>=75, df_new.loc[i, 'ES20rate']>85])])

            # ============email新增股票數量的文字============
            stock_numbers_all_template_text = f'<pre>策略股票數量 : \n\n\
T5(new):{len(new_T5_ID)} | T5_2(new):{len(new_T5_2_ID)} | T6(new):{len(new_T6_ID)} | T11(new):{len(new_T11_ID)} | all(new):{len(set(np.concatenate([new_T5_ID, new_T5_2_ID, new_T6_ID, new_T11_ID])))}\n\n\
T5(old):{len(old_T5_ID)} | T5_2(old):{len(old_T5_2_ID)} | T6(old):{len(old_T6_ID)} | T11(old):{len(old_T11_ID)} | all(old):{len(set(np.concatenate([old_T5_ID, old_T5_2_ID, old_T6_ID, old_T11_ID])))}\n\n\
T5(add):{len(T5_add_ID)} | T5_2(add):{len(T5_2_add_ID)} | T6(add):{len(T6_add_ID)} | T11(add):{len(T11_add_ID)} | all(add):{len(set(np.concatenate([T5_add_ID, T5_2_add_ID, T6_add_ID, T11_add_ID])))}\n\n\
T5(drop):{len(T5_drop_ID)} | T5_2(drop):{len(T5_2_drop_ID)} | T6(drop):{len(T6_drop_ID)} | T11(drop):{len(T11_drop_ID)} | all(drop):{len(set(np.concatenate([T5_drop_ID, T5_2_drop_ID, T6_drop_ID, T11_drop_ID])))}\n\n\</pre>'
            # print add and drop stock ID
            all_add_ID = np.array(list(set(np.concatenate([T5_add_ID, T5_2_add_ID, T6_add_ID, T11_add_ID]))))
            all_drop_ID = np.array(list(set(np.concatenate([T5_drop_ID, T5_2_drop_ID, T6_drop_ID]))))

            # ============email新增股票ID的文字============
            add_stockID_text = f'<pre>\n策略股票新增ID : \nT5新增股票ID : \n{T5_add_ID}\n\nT5RS破高刪除的股票ID : \n{T5_drop_goodID}\n\nT5_2新增股票ID : \n{T5_2_add_ID}\n\nT5_2RS破高刪除的股票ID : \n{T5_2_drop_goodID}\n\n\
T6新增股票ID : \n{T6_add_ID}\n\nT11新增股票ID : \n{T11_add_ID}\n</pre>'
            

            # ============email新增股票產業的表格文字============
            industry_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\產業別.xlsx').astype(int).astype(str)
            number_of_stock_in_industry = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\Stock_RS_rate_analysis\100產業分析\100產業RS排行.xlsx', index_col=0, header=0).loc['number'].astype(int)
            group_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\族群_複製.xlsx').astype(int).astype(str)
            number_of_stock_in_group = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\Stock_RS_rate_analysis\100產業分析\族群RS排行.xlsx', index_col=0, header=0).loc['number'].astype(int)
            concept_df = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\others\概念股_複製.xlsx').astype(int).astype(str)
            number_of_stock_in_concept = pd.read_excel(r'C:\Users\User\Desktop\StockInfoHub\Stock_RS_rate_analysis\100產業分析\概念股RS排行.xlsx', index_col=0, header=0).loc['number'].astype(int)
            industry_html_addstock_table, industry_html_dropstock_table = diff_strategy_stock(industry_df, all_add_ID, all_drop_ID, number_of_stock_in_industry)
            group_html_addstock_table, group_html_dropstock_table = diff_strategy_stock(group_df, all_add_ID, all_drop_ID, number_of_stock_in_group)
            concept_html_addstock_table, concept_html_dropstock_table = diff_strategy_stock(concept_df, all_add_ID, all_drop_ID, number_of_stock_in_concept)

            # ============email股票位階的文字============
            price_position_text = f'<pre>{new_date} ERS>80股票位階 : \n\
{numb_Price_higher_20MA}檔({str(round(100*numb_Price_higher_20MA/ers_higher80_length,2))}%)股票的收盤價大於20MA\n\
{numb_Price_higher_50MA}檔({str(round(100*numb_Price_higher_50MA/ers_higher80_length,2))}%)股票的收盤價大於50MA\n\
{numb_Price_higher_150MA}檔({str(round(100*numb_Price_higher_150MA/ers_higher80_length,2))}%)股票的收盤價大於150MA\n\
{numb_Price_higher_200MA}檔({str(round(100*numb_Price_higher_200MA/ers_higher80_length,2))}%)股票的收盤價大於200MA\n\n\
{new_date} 全部股票位階 : \n\
{all_numb_Price_higher_20MA}檔({str(round(100*all_numb_Price_higher_20MA/all_length,2))}%)股票的收盤價大於20MA\n\
{all_numb_Price_higher_50MA}檔({str(round(100*all_numb_Price_higher_50MA/all_length,2))}%)股票的收盤價大於50MA\n\
{all_numb_Price_higher_150MA}檔({str(round(100*all_numb_Price_higher_150MA/all_length,2))}%)股票的收盤價大於150MA\n\
{all_numb_Price_higher_200MA}檔({str(round(100*all_numb_Price_higher_200MA/all_length,2))}%)股票的收盤價大於200MA\n</pre>'
            
            # ============email全部文字============
            message_text = f'{new_date} vs {old_date}每日選股\n\n\
{stock_numbers_all_template_text}\n\
{add_stockID_text}\n\n'
            
            # ============產業類股ERS>80增減最多============
            ERS100 = pd.read_excel('C:/Users/User/Desktop/StockInfoHub/Stock_RS_rate_analysis/100產業分析/100產業RS排行.xlsx', header=0, index_col=0)
            yesterday = str(datetime.datetime.strptime(str(datetime.date.today() - datetime.timedelta(days=n_day_ago+delay_day)) , '%Y-%m-%d' ))
            for i in range(20):
                lastweek = str(datetime.datetime.strptime(str(datetime.date.today() - datetime.timedelta(days=n_day_ago+7+i)) , '%Y-%m-%d' ))
                try:
                    ERS100.loc[str(lastweek)]
                    print(f'lastweek: {lastweek}')
                    break
                except:
                    continue
            new_rs = ERS100.loc[str(day)]
            old_rs = ERS100.loc[str(yesterday)]
            lastweek_rs = ERS100.loc[str(lastweek)]
            add = new_rs - old_rs
            add.fillna(0, inplace=True)
            add = pd.concat([add, ERS100.loc['number']], axis=1)
            add.columns = ['add %', 'Num of stock in industry']
            add.sort_values(by = 'add %', ascending=False, inplace=True)
            first15_add = add.iloc[:15].transpose().to_html(index=True, header=True, justify='center')
            last15_add = add.iloc[-15:].transpose().iloc[:, ::-1].to_html(index=True, header=True, justify='center')
            lastweek_add = new_rs - lastweek_rs
            lastweek_add.fillna(0, inplace=True)
            lastweek_add = pd.concat([lastweek_add, ERS100.loc['number']], axis=1)
            lastweek_add.columns = ['weekly add %', 'Num of stock in industry']
            lastweek_add.sort_values(by = 'weekly add %', ascending=False, inplace=True)
            lastweek_first15_add = lastweek_add.iloc[:15].transpose().to_html(index=True, header=True, justify='center')
            lastweek_last15_add = lastweek_add.iloc[-15:].transpose().iloc[:, ::-1].to_html(index=True, header=True, justify='center')
            lastweek_add['count'] = lastweek_add['weekly add %']*lastweek_add['Num of stock in industry']/100
            lastweek_add.sort_values(by = 'count', ascending=False, inplace=True)
            lastweek_first15_add_count = lastweek_add.iloc[:15].drop('weekly add %', axis = 1).transpose().astype(int).to_html(index=True, header=True, justify='center')
            lastweek_last15_add_count = lastweek_add.iloc[-15:].drop('weekly add %', axis = 1).transpose().astype(int).iloc[:, ::-1].to_html(index=True, header=True, justify='center')
            message = """
            <html>
            <body>
                <p>{}</p>
                <p>{}</p>
                <p>新增策略股票產業別</p>
                {}
                <p>刪除策略股票產業別</p>
                {}
                <p>新增策略股票族群</p>
                {}
                <p>刪除策略股票族群</p>
                {}
                <p>新增策略概念股</p>
                {}
                <p>刪除策略概念股</p>
                {}
                <p>產業類股ERS>80今日增減%最多</p>
                {}
                {}
                <p>產業類股ERS>80一周內增減%最多</p>
                {}
                {}
                <p>產業類股ERS>80一周內增減數量最多</p>
                {}
                {}
            </body>
            </html>
            """.format(price_position_text, message_text
                       , industry_html_addstock_table, industry_html_dropstock_table
                       , group_html_addstock_table, group_html_dropstock_table
                       , concept_html_addstock_table, concept_html_dropstock_table
                       , first15_add, last15_add, lastweek_first15_add, lastweek_last15_add, lastweek_first15_add_count, lastweek_last15_add_count)
            #                <p>產業類股ERS>80一周內減少數量最多</p>
            # ============邮件服务器的信息 - 使用Gmail的SMTP服务器============
            smtp_server = 'smtp.gmail.com'
            smtp_port = 587  # Gmail的TLS端口号
            password = 'uqkz xwdm apft zmpr'  # 发件人的Gmail密码或应用程序专用密码
            
            # ============发件人和收件人信息============
            sender_email = 'johnnn1231232@gmail.com'  # 发件人的Gmail电子邮件地址
            if istest:
                receiver_email = ['johnnystockinfo@gmail.com']
            else:
                receiver_email = ['johnnystockinfo@gmail.com', 'jack12211221@gmail.com', 'kaiwenyang708@gmail.com']  # 收件人的Gmail电子邮件地址, 'jack12211221@gmail.com', 'kaiwenyang708@gmail.com', 'youping0116@gmail.com'
            
            # ============创建邮件对象============
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = ', '.join(receiver_email)
            if istest:
                msg['Subject'] = Header(f"{str(day).split(' ')[0]}-測試-每日選股", 'utf-8').encode()
            else:
                msg['Subject'] = Header(f"{str(day).split(' ')[0]}每日選股", 'utf-8').encode()
            msg.attach(MIMEText(message, 'html'))

            # ============打开要附加的xlsx文件============
            workbook = openpyxl.load_workbook(f'{path}{new_date}選股.xlsx')

            # ============保存工作簿为.xlsx文件============
            xlsx_filename = f'{new_date}選股.xlsx'
            workbook.save(xlsx_filename)

            # ============将.xlsx文件附加到邮件============
            with open(xlsx_filename, 'rb') as xlsx_file:
                excel_attachment = MIMEApplication(xlsx_file.read(), _subtype='xlsx')
                excel_attachment.add_header('Content-Disposition', 'attachment', filename=Header(xlsx_filename, 'utf-8').encode())
                msg.attach(excel_attachment)
            
            # ============连接到SMTP服务器============
            try:
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(sender_email, password)
                
                # ============发送邮件============
                server.sendmail(sender_email, receiver_email, msg.as_string())
                print(f'{bcolors.OK}寄送email OK{bcolors.RESET}')
                line_notify('sent email OK', TOKEN_FOR_UPDATE)
            except Exception as e:
                traceback.print_exc()
                print(f'\n{bcolors.FAIL}{e}{bcolors.RESET}')
                line_notify('⚠️sent email FAIL', TOKEN_FOR_UPDATE)
            finally:
                # ============关闭SMTP连接============
                server.quit()
            # ============删除附加的.xlsx文件============
            os.remove(xlsx_filename)
    except Exception as e:
        traceback.print_exc()
        print(f'\n{bcolors.FAIL}{e}{bcolors.RESET}')
        pass
line_notify(f'{day}daily_rs_industry.py執行完成', TOKEN_FOR_UPDATE)