# -*- codinD: utf-8 -*-
"""
Created on Tue Nov 14 02:46:22 2017

@author: Yao
"""

import price_slope_daily_new
import strategy_function
import daily_count_func
import report_func
import pandas as pd
import json
import requests
import datetime
import xlrd
import time
import random
import numpy as np
import math
import threading
import os

###在技術分析表中找尋股票名稱填入資訊
def excel_locate(xls_df_basic,stock):
    for i in range(len(xls_df_basic)):
        if xls_df_basic["股票名稱"][i] == stock:
            break
    return i

###計算每日報表用，需要將檔案交易日與個股交易日一致
def erase_for_big(xls_file_big_d,xls_df_report):
    xls_df_date = set(list(xls_df_report["日期"]))
    xls_df_big_date = set(list(xls_file_big_d["日期"]))
    diff = xls_df_big_date.difference(xls_df_date)
    diff = list(diff)   
    diff.sort(reverse = False)
    
    for i in range(xls_file_big_d.index[0],xls_file_big_d.index[-1]+1):
        if xls_file_big_d["日期"][i] in diff :
            xls_file_big_d.drop([i],axis = 0, inplace = True)
    xls_file_big_d.index = range(xls_df_report.index[-1]-len(xls_file_big_d)+1,xls_df_report.index[-1]+1)
    
    return xls_file_big_d

###尋找日期對應的周次
def be_format(data_date):
    year = data_date[:4]
    if data_date[4] == "0":
        month = data_date[5]
    else:
        month = data_date[4:6]
    if data_date[6] == "0":
        day = data_date[7]
    else:
        day = data_date[6:8]
    week = datetime.date(int(year),int(month),int(day)).isocalendar()
    return week
    
# 獲取西元時間(裝在list裡面)  
def GetDateList(j,xls_df):
    if DateGap(xls_df) is not None:
        DateList = []
        for x in range(j):
            TimePeriod = (datetime.datetime.now()-datetime.timedelta(days = x))
            TWyear = int(TimePeriod.strftime("%Y%m%d"))
            DateList.append(TWyear)
        return DateList
    else:
        return 0

# 讀取最近的日期 & 相差天數
def DateGap(xls_df):
    try:
        file_date = xls_df.iloc[0:1,0:1]
        file_date = str(file_date)[-8:]
        if file_date != "ndex: []":
            if file_date[4] == "0":
                month = file_date[5]
            else:
                month = file_date[4:6]
            if file_date[6] == "0":
                day = file_date[7]
            else:
                day = file_date[6:8]
            last_date = datetime.datetime(int(file_date[:4]),int(month),int(day))
            delta = str(datetime.datetime.now()-last_date)
            for k in range(len(delta)):
                if delta[k] == " ":
                    datedelta = delta[:k]
                    break  
            return int(datedelta)
    except:
        pass    

# 檢查是否有新股上市
def check_new_stock():    
    file = xlrd.open_workbook(r"D:\Stock Investment\股票分類-下載用(去除重複)-上市.xlsx")
    table = file.sheets()[0]

    date = str((str(datetime.datetime.now())[:4]+str(datetime.datetime.now())[5:7]+str(datetime.datetime.now())[8:10]))
    url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date="+str(date)+"&type=ALLBUT0999&_=1562075405673"
    headers = {"User-Agent":"Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.167 Safari/537.36",
               "Cache-Control":"max-age=0",
               "Accept-Language":"zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
               "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
               "Accept-Encoding":"gzip, deflate",
               "Connection":"keep-alive",
               "Host":"www.twse.com.tw",
               "Upgrade-Insecure-Requests":"1"}
    sourcecode = requests.get(url, headers = headers)
    rawdata = sourcecode.text
    time.sleep(random.randint(3,8))
    no_stock_list = []

    if "data9" in rawdata:
        data_want = json.loads(rawdata)["data9"]
        for data in data_want:
            break_switch = 0
            if data[0].strip()[:2] == "00" or len(data[0].strip()) > 4 or "-DR" in data[1].strip():
                continue
            for i in range(num_stock):
                if break_switch == 1:
                    break
                stock_list = table.row_values(i)[2:]
                for stock in stock_list:
                    if stock != '' and "*" in data[1].strip() and stock == data[1].strip().replace("*",""):
                        break_switch = 1
                        break
                    elif stock != '' and "*" not in data[1].strip() and stock == data[1].strip() :
                        break_switch = 1
                        break
            if break_switch == 0:
                no_stock_list.append(data[1].strip())
                print(no_stock_list)
    
        if len(no_stock_list) != 0 :
            no_stock_list = pd.DataFrame(no_stock_list, columns = ["新股清單"] ) 
            no_stock_list.to_excel("D:\\Stock Investment\\新股清單-市"+date+".xlsx")     

###計算當日推薦個股報表
def daily_category_count(): 
    xls_file_cate = pd.ExcelFile("D:\\Stock Investment\\個股分類表.xlsx")
    xls_df_cate = xls_file_cate.parse(xls_file_cate.sheet_names[0], index_col=[0])
        
    file_industry = xlrd.open_workbook("D:\\Stock Investment\\股票分類-下載用(去除重複)-分類.xlsx")
    table_industry = file_industry.sheets()[0]
       
    switch = 0
    while switch == 0:
        if os.path.isfile(r"C:\\Users\\User\\Desktop\\每日計算報告\\上市總表"+str(today_std)+".xls"):
            switch = 1        
        else:
            time.sleep(120)
            
    xls_file_basic = pd.ExcelFile(r"C:\\Users\\User\\Desktop\\每日計算報告\\上市總表"+str(today_std)+".xls")    
    xls_df_basic = xls_file_basic.parse(xls_file_basic.sheet_names[0], index_col=[0])
      
    candidate_list= []
    for i in range(len(xls_df_basic)):
        switch = 0 
        if xls_df_basic["日價格日布林 %B"][i] >= 100:
            if xls_df_basic["訊號-60日內"][i] == "V" : 
                switch = 1
            elif xls_df_basic["訊號-今日"][i] == "V" :
                switch = 1  
        if type(xls_df_basic["訊號發動盤堅天數"][i]) == str :  
            switch = 1 
        
        if switch == 1:
            candidate_list.append(xls_df_basic["股票名稱"][i])
################################################     
#### 主要運行程式碼-Daily Report-產業/類股籌碼分析 #
###############################################  
    proceed,cycle_day = 0,59
    while proceed == 0 :        
        xls_file_switch = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
        xls_df_switch = xls_file_switch.parse(xls_file_switch.sheet_names[0], index_col=[0])
        if xls_df_switch["計算"][0] == 1 :
            proceed = 1
        else:
            time.sleep(60)
            continue 
                                            
    for i in range(num_stock):
        stock_list = table.row_values(i)[2:]
        for stock in stock_list:
            if stock != '' and stock in candidate_list:    
                switch = 0
                category = table.row_values(i)[1]
                if "工業" in category:
                    loca = category.strip().index("工")
                    category = category.strip()[:loca]
                elif "業" in category and category != "航運業" :
                    loca = category.strip().index("業")
                    category = category.strip()[:loca]
    
                xls_df_report = pd.read_csv("C:\\Users\\User\Desktop\\Daily Report\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)
              
                weight_list = []           
                for j in range(len(xls_df_cate)):
                    if xls_df_cate["股票名稱"][j] == stock :
                        try:
                            if math.isnan(xls_df_cate["分類"][j]) == True:
                                break
                        except:
                            weight_list = list(set(xls_df_cate["分類"][j].split(",")))                    
                        break
    
                if len(weight_list) == 0 :
                    continue
           
                date_list = [] 
                for k in range(len(weight_list)):
                    if date_list != [] :
                        break
                    for j in range(1,len(xls_df_report["日期"])):                  
                        try:
                            if float(xls_df_report[weight_list[k]+"股數"][len(xls_df_report["日期"])-j]) == 0.0001:
                                date_list.append(xls_df_report["日期"][len(xls_df_report["日期"])-j])
                            else:
                                break
                        except:
                            continue
          
                if len(date_list) == 0:
                    continue    

                xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤.xlsx")
                xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])    
                xls_df_big = erase_for_big(xls_file_big_d,xls_df_report)
                    
                date_list.sort()
                for data_date in date_list  :                         
                    for k in range(len(xls_df_report["日期"])-60,len(xls_df_report["日期"])):
                        if int(xls_df_report["日期"][k]) == int(data_date):
                            break
    
                    category_3 = xls_df_report["股數比"][k] 
                    for weight in weight_list :
                        try:
                            category_money_tracking = daily_count_func.category_money_tracking(xls_df_report,weight,cycle_day,category,k,xls_df_big,category_3)           
                            xls_df_report["股數比"][k] = category_money_tracking[0] 
                            xls_df_report["股數值"][k] = category_money_tracking[1] 
                            xls_df_report["股數值_1"][k] = category_money_tracking[2]
                            xls_df_report["股數"][k] = category_money_tracking[3]
                            xls_df_report["金額"][k] = category_money_tracking[10]
                            xls_df_report["股數增減率"][k] = category_money_tracking[4]
                            xls_df_report["股數_1"][k] = category_money_tracking[5]
                            xls_df_report["股數值_1"][k] = category_money_tracking[6]
                            xls_df_report["金額_1"][k] = category_money_tracking[7]
                            xls_df_report["金額值_1"][k] = category_money_tracking[8]
                            xls_df_report["金額_2"][k] = category_money_tracking[9]
                            switch = 1
                        except:
                            print(stock+"沒有"+weight)
                            print(str(datetime.datetime.now())) 
            
                if switch == 1:
                    xls_df_report.index = range(len(xls_df_report)) 
                    xls_df_report.to_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv", encoding = "big5")                            
                    print(stock+"日報表-產業/類股籌碼分析 完成")
                    print(str(datetime.datetime.now())) 
                else:
                    print(stock+"日報表-產業/類股籌碼分析 不需變動")
                    print(str(datetime.datetime.now()))    
    
###############################
# 主要運行程式碼-計算candidate #
#############################
    for i in range(num_stock):
        stock_list = table.row_values(i)[2:]
        for stock in stock_list:
            if stock != '' and stock in candidate_list:     
                category = table.row_values(i)[1]

                if "工業" in category:
                    loca = category.strip().index("工")
                    category = category.strip()[:loca]
                elif "業" in category and category != "觀光事業" and category != "航運業" :
                    loca = category.strip().index("業")
                    category = category.strip()[:loca]
                    
                xls_df_report = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)
                
                xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤_計算用.xlsx")
                xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])    
                xls_df_big = erase_for_big(xls_file_big_d,xls_df_report)
                               
                locate = excel_locate(xls_df_basic,stock)
                print(stock+"ready to count")                
                price_slope_daily_new.industry_stock(stock,xls_df_report,xls_df_basic,xls_df_cate,table_industry,table_category,locate,category,xls_df_big)
                print(stock+"已經完成計算")
                print(str(datetime.datetime.now()))                              
        
    xls_df_basic.to_excel(r"C:\\Users\\User\\Desktop\\每日計算報告\\上市總表"+str(today_std)+".xls")  

    xls_file_basic = pd.ExcelFile(r"C:\\Users\\User\\Desktop\\每日計算報告\\上市總表"+str(today_std)+".xls")    
    xls_df_basic = xls_file_basic.parse(xls_file_basic.sheet_names[0], index_col=[0])                         
################################
## 主要運行程式碼-產出強勢股報表最終報表 #
##############################    
    strategy_function.report_count(xls_df_basic,today_std,table_category,delete_date,switch_big,yesterday_std)
################################################     
### 主要運行程式碼-Daily Report-產業/類股籌碼分析 #
###############################################   
    for i in range(num_stock):
        stock_list = table.row_values(i)[2:]
        for stock in stock_list:
            if stock != ''  and stock not in no_need :
                switch = 0
                category = table.row_values(i)[1]
                if "工業" in category:
                    loca = category.strip().index("工")
                    category = category.strip()[:loca]
                elif "業" in category and category != "觀光事業" and category != "航運業" :
                    loca = category.strip().index("業")
                    category = category.strip()[:loca]
    
                xls_df_report = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)
                         
                for j in range(len(xls_df_cate)):
                    if xls_df_cate["股票名稱"][j] == stock :
                        try:
                            if math.isnan(xls_df_cate["分類"][j]) == True:
                                break
                        except:
                            weight_list = list(set(xls_df_cate["分類"][j].split(",")))                    
                        break
        
                if len(weight_list) == 0 :
                    continue
           
                date_list = [] 
                for k in range(len(weight_list)):
                    if date_list != [] :
                        break
                    for j in range(1,len(xls_df_report["日期"])):
                        try:
                            if float(xls_df_report[weight_list[k]+"股數"][len(xls_df_report["日期"])-j]) == 0.0001:
                                date_list.append(xls_df_report["日期"][len(xls_df_report["日期"])-j])
                            else:
                                break
                        except:
                            continue
          
                if len(date_list) == 0:
                    continue    
 
                xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤_計算用.xlsx")
                xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])    
                xls_df_big = erase_for_big(xls_file_big_d,xls_df_report)
                   
                date_list.sort()
                for data_date in date_list  :                         
                    for k in range(len(xls_df_report["日期"])-60,len(xls_df_report["日期"])):
                        if int(xls_df_report["日期"][k]) == int(data_date):
                            break
    
                    category_3 = xls_df_report["成交"][k] 
                    for weight in weight_list :
                        try:
                            category_money_tracking = daily_count_func.category_money_tracking(xls_df_report,weight,cycle_day,category,k,xls_df_big,category_3)           
                            xls_df_report["股數比"][k] = category_money_tracking[0] 
                            xls_df_report["股數值"][k] = category_money_tracking[1] 
                            xls_df_report["股數值_1"][k] = category_money_tracking[2]
                            xls_df_report["股數"][k] = category_money_tracking[3]
                            xls_df_report["金額"][k] = category_money_tracking[10]
                            xls_df_report["股數增減率"][k] = category_money_tracking[4]
                            xls_df_report["股數_1"][k] = category_money_tracking[5]
                            xls_df_report["股數值_1"][k] = category_money_tracking[6]
                            xls_df_report["金額_1"][k] = category_money_tracking[7]
                            xls_df_report["金額值_1"][k] = category_money_tracking[8]
                            xls_df_report["金額_2"][k] = category_money_tracking[9]
                            switch = 1
                        except:
                            print(stock+"沒有"+weight)
                            print(str(datetime.datetime.now())) 
            
                if switch == 1:
                    xls_df_report.index = range(len(xls_df_report)) 
                    xls_df_report.to_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv", encoding = "big5")                            
                    print(stock+"日報表-分析 完成")
                    print(str(datetime.datetime.now())) 
                else:
                    print(stock+"日報表-分析 不需變動")
                    print(str(datetime.datetime.now())) 
    
    xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0])
    xls_df["計算"][0] = 0
    print("計算=0")
    xls_df.to_excel("D:\\Stock Investment\\個股計算開關-市.xlsx")                      
            
def xls_df_price_():
    
    xls_file = pd.ExcelFile("D:\\Stock Investment\\漲跌幅報表-上市-2022.xlsx")
    xls_df_price = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
        
    columns = xls_df_price.columns
    time_period = GetDateList(DateGap(xls_df_price),xls_df_price)                
    for date in time_period :
        today = datetime.datetime(int(str(date)[:4]),int(str(date)[4:6]),int(str(date)[6:]))
        if today.weekday() == 6 or today.weekday() == 5 or date in skip_date:
            continue
        
        list_need = [today_std]
        for k in range(len(xls_df_price.columns)-1):  
            list_need.append("")
           
    try:       
        data = pd.DataFrame([list_need] ,columns = columns)
        below = xls_df_price.loc[:]
        xls_df_price = data.append(below,ignore_index=True)  
    except:
        pass

    return xls_df_price

                   
switch = 0
while switch == 0:
    if  int(str(datetime.datetime.now())[11:13]+str(datetime.datetime.now())[14:16]) >= 1435 :    
        switch = 1
    else:
        time.sleep(300) 
  
#上市    
file = xlrd.open_workbook(r"D:\Stock Investment\股票分類-下載用(去除重複)-上市.xlsx")
table = file.sheets()[0]
sheet = file.sheets()[1]

file_category = pd.ExcelFile("D:\\Stock Investment\\股票分類-上市類.xlsx")
table_category = file_category.parse(file_category.sheet_names[0])

xls_file_multiply = pd.ExcelFile("D:\\Stock Investment\\估表.xlsx")
xls_df_multiply = xls_file_multiply.parse(xls_file_multiply.sheet_names[0], index_col=[0])

xls_file_basic = pd.ExcelFile("D:\\Stock Investment\\技術分析表-上市.xlsx")
xls_df_basic = xls_file_basic.parse(xls_file_basic.sheet_names[0],index_col=0)

no_need = []

skip_date = []

headers = {"User-Agent":"Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, lik1268e Gecko) Chrome/64.0.3282.167 Safari/537.36",
           "Cache-Control":"max-age=0",
           "Accept-Language":"zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7", 
           "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
           "Accept-Encoding":"gzip, deflate",
           "Connection":"keep-alive",
           "Host":"www.twse.com.tw",
           "Upgrade-Insecure-Requests":"1"}

enter,run_again,num_stock,switch_report,switch_big = 0,7,33,1,0

while True:

    today_std = int((str(datetime.datetime.now())[:4]+str(datetime.datetime.now())[5:7]+str(datetime.datetime.now())[8:10]))

    if run_again == 1:
        xls_df_price = xls_df_price_()
        delete_date = int(xls_df_price["日期"][15])
        yesterday_std = int(xls_df_price["日期"][1])
#############################
# 主要運行程式碼-類股金額 ###
############################          
    while run_again == 1:       
        fail_time = 0
        xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤.xlsx")
        xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
        columns = xls_df.columns
        fail,data_all = [],[]        
        if GetDateList(DateGap(xls_df),xls_df) == 0:
            continue               # 決定資料天數  
        time_period = GetDateList(DateGap(xls_df),xls_df) 

        for date in time_period :
            today = datetime.datetime(int(str(date)[:4]),int(str(date)[4:6]),int(str(date)[6:]))
            if today.weekday() == 6 or today.weekday() == 5 or date in skip_date:
                continue 
        
            data_single = []
            for k in range(len(xls_df.columns)):
                if k == 0:
                    data_single.append(date)
                else:
                    data_single.append("")
                
            data_single = pd.DataFrame([data_single], columns = columns )       
        
            try:
                url = "https://www.twse.com.tw/rwd/zh/afterTrading/BFIAMU?response=json&date="+str(date)+"&_=1547953498712"
                sourcecode = requests.get(url, headers = headers)
                rawdata = sourcecode.text
                time.sleep(random.randint(3,8))
    
                if "data" in rawdata:
                    data_want = json.loads(rawdata)["data"]
                    for data in data_want:
                        for title in xls_df.columns:                                
                            if data[0].strip() == title:
                                loca = data[0].strip().index("類")
                                key = data[0].strip()[:loca]
                                for k in [1,2,3,4]:
                                    try:
                                        data[k] = float(data[k])
                                    except:
                                        try:
                                            data[k] = int(float((data[k].replace(",",""))))
                                        except:
                                            if type(data[k]) == str :
                                                continue
    
                                data_single[key+"成交股數"][0] = int(data[1])
                                data_single[key+"成交金額"][0] = int(data[2])
                                data_single[key+"成交筆數"][0] = int(data[3])
                                data_single[key+"漲跌指數"][0] = data[4] 
                            
                    data_all.append(np.array(data_single).tolist()[0])                            
                    print(str(date)+"類股金額完成")  
            except:
                fail.append(key)
                continue
    
        if len(fail) != 0 or len(data_all) == 0:
            time.sleep(300)
            continue                       
        data_all = pd.DataFrame(data_all, columns = columns )      
        below = xls_df.loc[:]
        xls_df = data_all.append(below,ignore_index=True)
     
        for date in time_period:
            for i in range(len(xls_df["日期"])):
                if int(xls_df["日期"][i]) == int(date):
                    for j in range(len(xls_df.columns)):
                        try:
                            if math.isnan(xls_df[xls_df.columns[j]][i]): 
                                if "類指數" in xls_df.columns[j]:
                                    continue
                                else:
                                    fail_time = fail_time + 1
                                    print("指數沒更新到最新") 
                                
                        except:
                            continue
    
        if fail_time == 0 :
            run_again = 2 
        elif fail_time == 1 :
            run_again = 1
        elif fail_time == 2 :
            run_again = 2              
    
#############################
# 主要運行程式碼-類股指數 ###
############################
    while run_again == 2:
        fail_time = 0
        for date in time_period:  
            for i in range(len(time_period)):
                if int(xls_df["日期"][i]) == int(date):
                    try:
                        url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date="+str(xls_df["日期"][i])+"&type=IND&_=1563720943613"
                        sourcecode = requests.get(url, headers = headers)
                        rawdata = sourcecode.text
                        time.sleep(random.randint(3,8))
                        if "data1" in rawdata:
                            data_want = json.loads(rawdata)["data1"]
                            for data in data_want:
                                for title in xls_df.columns:    
                                    if data[0] == "航運類指數":
                                        data[0] = "航運業類指數"
                                    elif data[0] == "電子工業類指數":
                                        data[0] = "電子類指數"                                        
                                    elif data[0] == "發行量加權股價指數":
                                        continue
                                    if data[0].strip() in title:
                                        loca = data[0].strip().index("類")
                                        key = data[0].strip()[:loca]
                                        for k in [1,4]:
                                            try:
                                                data[k] = float(data[k])
                                            except:
                                                try:
                                                    data[k] = int(float((data[k].replace(",",""))))
                                                except:
                                                    if type(data[k]) == str :
                                                        continue
                                        xls_df[key+"類指數"][i] = data[1]
                                        xls_df[key+"漲跌百分比(%)"][i] = data[4]
                                        break
                                
                        if "data1" in rawdata:
                            data_want = json.loads(rawdata)["data1"]
                            for data in data_want:
                                if data[0] == "發行量加權股價指數" :
                                    for k in [1,3]:
                                        try:
                                            data[k] = float(data[k])
                                        except:
                                            try:
                                                data[k] = float((data[k].replace(",","")))
                                            except:
                                                if type(data[k]) == str :
                                                    continue
    
                                    xls_df["發行量加權股價指數"][i] = float(data[1])
                                    xls_df["漲跌點數"][i] = float(data[3])
                                    break
    
                    except:
                        print("下載失敗，須注意"+str(datetime.datetime.now()))
                        xls_df.to_excel("D:\\Stock Investment\\股票數據\\市盤.xlsx")
                        time.sleep(10)
    
            print(str(date)+"類股指數完成")
       
        for date in time_period:
            for i in range(len(xls_df["日期"])):
                if int(xls_df["日期"][i]) == int(date):
                    for j in range(len(xls_df.columns)):
                        try:
                            if math.isnan(xls_df[xls_df.columns[j]][i]): 
                                fail_time = fail_time + 1
                                print("指數沒更新到最新") 
                                
                        except:
                            continue
    
        if fail_time == 0 :
            run_again = 3
        elif fail_time == 1 :
            run_again = 2 
        elif fail_time == 2 :
            run_again = 3
    
    #############################
    # 主要運行程式碼-大盤指數 ###
    ############################
    while run_again == 3:
        fail_time = 0
        for date in time_period:  
            for i in range(len(time_period)):
                if int(xls_df["日期"][i]) == int(date):
                    try:
                        url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date="+str(xls_df["日期"][i])+"&_=1564278168508"
                        sourcecode = requests.get(url, headers = headers)
                        rawdata = sourcecode.text
                        time.sleep(random.randint(3,8))                    
                        if "data7" in rawdata:
                            data_want = json.loads(rawdata)["data7"]
                            for data in data_want:
                                if data[0] == "1.一般股票" :
                                    for k in [1,2,3]:
                                        try:
                                            data[k] = float(data[k])
                                        except:
                                            try:
                                                data[k] = int(float((data[k].replace(",",""))))
                                            except:
                                                if type(data[k]) == str :
                                                    continue
                                    xls_df["成交金額"][i] = float(data[1])
                                    xls_df["成交股數"][i] = float(data[2])
                                    xls_df["成交筆數"][i] = float(data[3])
                                    break
    
                    except:
                        print("下載失敗，須注意"+str(datetime.datetime.now()))
                        xls_df.to_excel("D:\\Stock Investment\\股票數據\\市盤.xlsx")
                        time.sleep(10)
    
            print(str(date)+"類股指數完成")
      
        for date in time_period:
            for i in range(len(xls_df["日期"])):
                if int(xls_df["日期"][i]) == int(date):
                    for j in range(len(xls_df.columns)):
                        try:
                            if math.isnan(xls_df[xls_df.columns[j]][i]): 
                                fail_time = fail_time + 1
                                print("指數沒更新到最新") 
                                
                        except:
                            continue
    
        if fail_time == 0 :
            run_again = 4
        elif fail_time == 1 :
            run_again = 3 
        elif fail_time == 2 :
            run_again = 4
        
        xls_df.to_excel("D:\\Stock Investment\\股票數據\\市盤.xlsx")
        print("類股指數下載完成"+str(datetime.datetime.now()))

###########################################
# 主要運行程式碼-成為周制-每日更新#   
##########################################
        xls_df_big_d = xls_df
        xls_df_big_d = xls_df_big_d.sort_values(by = "日期",ascending = True)  
        xls_df_big_d.index = range(len(xls_df_big_d))

        xls_file_big_w = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤_周.xlsx")
        xls_df_big_w = xls_file_big_w.parse(xls_file_big_w.sheet_names[0], index_col=[0])            
        
        today = int(xls_df_big_d["日期"][len(xls_df_big_d)-1])
        last_day = int(xls_df_big_w["日期"][len(xls_df_big_w)-1])
        
        p = 2
        if be_format(str(today))[1] == be_format(str(last_day))[1]  :
            xls_df_big_w = xls_df_big_w.drop([len(xls_df_big_w)-1]) 
            p = 1
        
        lastday_list = []
        switch = 0
        
        cycle_week = 11
        
        for column in xls_df_big_d.columns:
            stock_sum,money_sum,trans_sum = 0,0,0
            if "成交股數" in column and "比重" not in column and "%B(季)" not in column:        
                for j in range(len(column)):
                    if column[j] == "成":
                        category = column[:j]
                        break
                for i in range(len(xls_df_big_d)-14,len(xls_df_big_d)):
                    if int(xls_df_big_d["日期"][i]) <= int(xls_df_big_w["日期"][len(xls_df_big_w)-p]) :
                        continue                       
                    if be_format(str(xls_df_big_d["日期"][i]))[1] == be_format(str(xls_df_big_d["日期"][i-1]))[1] or stock_sum == 0 :
                        stock_sum = xls_df_big_d[column][i] + stock_sum
                        money_sum = xls_df_big_d[category+"成交金額"][i] + money_sum
                        trans_sum = xls_df_big_d[category+"成交筆數"][i] + trans_sum
                    else:
                        xls_df_big_d[column][i-1] = stock_sum
                        xls_df_big_d[category+"成交金額"][i-1] = money_sum   
                        xls_df_big_d[category+"成交筆數"][i-1] = trans_sum  
                        stock_sum,money_sum,trans_sum = 0,0,0
                        stock_sum = xls_df_big_d[column][i] + stock_sum
                        money_sum = xls_df_big_d[category+"成交金額"][i] + money_sum 
                        trans_sum = xls_df_big_d[category+"成交筆數"][i] + trans_sum
                        if switch == 0 :
                            lastday_list.append(i-1)
                    if i == len(xls_df_big_d)-1 and be_format(str(xls_df_big_d["日期"][i]))[1] == be_format(str(xls_df_big_d["日期"][i-1]))[1]:
                        xls_df_big_d[column][i] = stock_sum
                        xls_df_big_d[category+"成交金額"][i] = money_sum    
                        xls_df_big_d[category+"成交筆數"][i] = trans_sum
                        if switch == 0 :
                            lastday_list.append(i)
                            
                switch = 1        
        
        add_column = xls_df_big_d.iloc[len(xls_df_big_d)-1:len(xls_df_big_d),:]
        xls_df_big_w = xls_df_big_w.append(add_column)
        xls_df_big_w.index = range(len(xls_df_big_w))      
        
        xls_df_big_w.to_excel("D:\\Stock Investment\\股票數據\\市盤_周.xlsx")      
        xls_df_big_w.iloc[len(xls_df_big_w)-40:].to_excel("D:\\Stock Investment\\股票數據\\市盤_周_計算用.xlsx")            
        print("類股指數-周-每日更新下載完成"+str(datetime.datetime.now()))

##############################
# 主要運行程式碼-每日收盤行情 #
#############################  
    fail_1 = fail_2 = []            
    while run_again == 4:
        column_title =['日期',"證券代號","證券名稱","成交股數","成交筆數","成交金額","開盤價","最高價","最低價","收盤價","漲跌(+/-)","漲跌價差","最後揭示買價","最後揭示買量","最後揭示賣價","最後揭示賣量","本益比"]
        column_title_new =['日期',"成交股數","成交筆數","成交金額","開盤價","最高價","最低價","收盤價","漲跌價差","最後揭示買價","最後揭示買量","最後揭示賣價","最後揭示賣量","本益比"]  
        for i in range(num_stock):
            stock_list = table.row_values(i)[2:]
            number_list = sheet.row_values(i)[2:]
            for stock in stock_list:
                data_list = []
                if stock != '' : 
                    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-每日行情.xls")
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0])  
                    if GetDateList(DateGap(xls_df),xls_df) == 0:
                        continue               # 決定資料天數  
                    time_period = GetDateList(DateGap(xls_df),xls_df) 

                    try:
                        for date in time_period :
                            today = datetime.datetime(int(str(date)[:4]),int(str(date)[4:6]),int(str(date)[6:]))
                            if today.weekday() == 6 or today.weekday() == 5 or date in skip_date:
                                continue
    
                            url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date="+str(date)+"&type=ALLBUT0999&_=1562075405673"
                            sourcecode = requests.get(url, headers = headers)
                            rawdata = sourcecode.text
                            if switch_report == 0 :
                                time.sleep(random.randint(5,8))  
                            else:
                                time.sleep(random.randint(3,5))
                                
                            if "data9" in rawdata:
                                data_want = json.loads(rawdata)["data9"]
                                for data in data_want:
                                    if data[0].strip() == str(int(number_list[stock_list.index(stock)])): 
                                        data.insert(0,int(date))
                                    
                                        for k in [3,4,5,6,7,8,9]:
                                            if k == 3 or k == 4 or k == 5 or k == 6 or k == 7 or k == 8 or k == 9:
                                                try:
                                                    data[k] = float(data[k])
                                                except:
                                                    try:
                                                        data[k] = int(float((data[k].replace(",",""))))
                                                    except:
                                                        if type(data[k]) == str :
                                                            continue                                        
                                                                            
                                        data_list.append(data)
                                        break
                    
                        if len(data_list) != 0 :
                            info = pd.DataFrame(data_list,columns = column_title )
                            del info["證券代號"]
                            del info['證券名稱']
                            del info["漲跌(+/-)"]
                            below = xls_df.loc[:]
                            below.columns = column_title_new
                            newdata = info.append(below,ignore_index=True)
                            daily_count_func.up_down_day(stock,newdata,xls_df_price)
                            ##計算報告
                            if switch_report == 1 and stock not in no_need:
                                report_func.daily_count(newdata,stock,xls_df_multiply,switch_big) 
                                report_func.week_count_and_tdcc(table,i,stock,switch_big)
                                report_func.daily_power_count(stock,category,switch_big)
                                report_func.daily_category_count(stock,xls_df_basic,no_need,switch_big,category)
                            newdata.to_excel("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-每日行情.xls")  # 匯出excel
                            print(stock+"每日收盤行情已經下載")
                            print(str(datetime.datetime.now()))
                            newdata = 0  #防錯機制，假設20180407有資料，因為沒抓到而改用20180408，存檔會變成上一個檔案的資料
                                                                                                      
                    except:
                        print(stock+"每日收盤行情下載失敗，須注意")
                        print(str(datetime.datetime.now()))
                        continue
           
    
    ###檢查是否有更新到今日日期                             
        n2 = 0
        fail_list = []    
        for j in range(num_stock):
            stock_list = table.row_values(j)[2:]
            number_list = sheet.row_values(j)[2:]
            for stock in stock_list:            
                if stock != '' :
                    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(j)[1]+"\\"+stock+"\\"+stock+"-每日行情.xls")
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                    last_date = xls_df["日期"][0]
                    if last_date == int(str(datetime.datetime.now())[:4]+str(datetime.datetime.now())[5:7]+str(datetime.datetime.now())[8:10]):
                        continue
                    else:
                        fail_list.append(stock)
                        print(stock+"沒更新到最新")
                        
        if fail_list == []:
            run_again = 5
            pass
        elif fail_1 == [] :
            fail_1 = fail_list
            continue
        elif fail_2 == [] :
            fail_2 = fail_list
            n2 = 1

        if n2 == 1:
            if len(fail_1) == len(fail_2):
                for i in range(len(fail_1)):
                    if fail_1[i] == fail_2[i]:
                        run_again = 5
                        continue
                    else:
                        run_again = 4
                        fail_1,fail_2 = [],[]
                        n2 = 0
                        break
            else:                
                run_again = 4
                fail_1,fail_2 = [],[]
                n2 = 0
                continue
                        
        for column in xls_df_price.columns:
            if xls_df_price[column][0] == "":
                xls_df_price[column][0] = 0
        xls_df_price.to_excel("D:\\Stock Investment\\漲跌幅報表-上市-2022.xlsx")
        xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
        xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0])         
        xls_df = xls_df.drop(["失敗清單"], axis=1)
        fail_stock = pd.DataFrame(list(set(fail_list)) ,columns = ["失敗清單"] )
        xls_df = pd.concat([xls_df,fail_stock],axis = 1)
        xls_df["每日行情"][0] = 1
        xls_df.to_excel("D:\\Stock Investment\\個股計算開關-市.xlsx")
        print("都有更新到最新")
     

###################################
# 主要運行程式碼-分析 #
#################################
    if run_again == 5:         
        thread_3 = threading.Thread(target=check_new_stock)
        thread_3.start()          
        thread_5 = threading.Thread(target=daily_category_count)
        thread_5.start()   

##############################
# 主要運行程式碼-三大法人 #
#############################
    fail_1 = fail_2 = []   
    switch_ = 0
    while run_again == 6:
        column_title =["日期","證券代號",'證券名稱',"外陸資買進股數(不含外資自營商)","外陸資賣出股數(不含外資自營商)","外陸資買賣超股數(不含外資自營商)","外資自營商買進股數",
                       "外資自營商賣出股數","外資自營商買賣超股數","投信買進股數","投信賣出股數","投信買賣超股數","自營商買賣超股數","自營商買進股數(自行買賣)",
                       "自營商賣出股數(自行買賣)","自營商買賣超股數(自行買賣)","自營商買進股數(避險)","自營商賣出股數(避險)","自營商買賣超股數(避險)","三大法人買賣超股數"]
    
        column_title_new =["日期","外資買進股數","外資賣出股數","外資買賣超股數","投信買進股數","投信賣出股數","投信買賣超股數","自營商買賣超股數","自營商買進股數(自行買賣)",
                               "自營商賣出股數(自行買賣)","自營商買賣超股數(自行買賣)","自營商買進股數(避險)","自營商賣出股數(避險)","自營商買賣超股數(避險)","三大法人買賣超股數"]

        for i in range(num_stock):
            stock_list = table.row_values(i)[2:]
            number_list = sheet.row_values(i)[2:]
            for stock in stock_list:
                data_list = []
                if stock != '':  
                    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-三大法人買賣超.xls")
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0])                    
                    xls_file_o = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-比率.xls")
                    xls_df_o = xls_file_o.parse(xls_file_o.sheet_names[0], index_col=[0]) 
                    category = table.row_values(i)[1] 
                    if GetDateList(DateGap(xls_df),xls_df) == 0:
                        continue               # 決定資料天數  
                    time_period = GetDateList(DateGap(xls_df),xls_df) 

                    try:
                        for date in time_period : 
                            if date not in np.array(xls_df_o["日期"]):
                                continue
                            pass_ = 0
                            today = datetime.datetime(int(str(date)[:4]),int(str(date)[4:6]),int(str(date)[6:]))
                            if today.weekday() == 6 or today.weekday() == 5 or date in skip_date:
                                time_period.remove(date)
                                continue
    
                            url = "https://www.twse.com.tw/rwd/zh/fund/T86?response=json&date="+str(date)+"&selectType=ALLBUT0999&_=1562078116855"
                            sourcecode = requests.get(url, headers = headers)
                            rawdata = sourcecode.text
                            time.sleep(random.randint(5,8))   
                                
                            if "200" in str(sourcecode):
                                switch = 2
                            else:
                                switch = 0
                            if "data" in rawdata:
                                data_want = json.loads(rawdata)["data"]
                                pass_ = 1
                                for data in data_want:
                                    if data[0].strip() == str(int(number_list[stock_list.index(stock)])): 
                                        data.insert(0,str(date))
                                        
                                        for k in [3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19]:                       
                                            try:
                                                data[k] = float(data[k])
                                            except:
                                                data[k] = int((data[k]).replace(",",""))
                                        
                                        data_list.append(data)
                                        switch = 1
                                        break
                                    
                                if switch == 2 and pass_ == 1 :
                                    if date in np.array(xls_df_o["日期"]):
                                        data_list.append([date,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0])                                
                            if switch == 0 : 
                                if date in np.array(xls_df_o["日期"]):
                                    data_list.append([date,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0])                              
    
                        if len(data_list) != 0 :
                            info = pd.DataFrame(data_list,columns = column_title )
                            del info["證券代號"]
                            del info['證券名稱']
                            del info["外資自營商買進股數"]
                            del info['外資自營商賣出股數']
                            del info["外資自營商買賣超股數"]
                            info.columns = column_title_new
                            below = xls_df.loc[:]
                            newdata = info.append(below,ignore_index=True)
                            newdata.to_excel("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-三大法人買賣超.xls")  # 匯出excel
                            newdata = 0  #防錯機制，假設20180407有資料，因為沒抓到而改用20180408，存檔會變成上一個檔案的資料
                            print(stock+"三大法人買賣超已經下載")
                            print(str(datetime.datetime.now()))
    
                        if switch_ == 0 :
                            today_std = int(info["日期"][0])
                            xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
                            xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                            fail_stock = np.array(xls_df["失敗清單"]).tolist()
                            fail_stock = [x for x in fail_stock if str(x)!= 'nan']
                            switch_ = 1                            
               
                    except:
                        print(stock+"三大法人買賣超下載失敗，須注意")
                        print(str(datetime.datetime.now()))
                        continue
               
    ###檢查是否有更新到今日日期                    
        n2 = 0
        fail_list = []    
        for j in range(num_stock):
            stock_list = table.row_values(j)[2:]
            number_list = sheet.row_values(j)[2:]
            for stock in stock_list:            
                if stock != '':
                    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(j)[1]+"\\"+stock+"\\"+stock+"-三大法人買賣超.xls")
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                    last_date = int(xls_df["日期"][0])
                    if last_date == int(str(datetime.datetime.now())[:4]+str(datetime.datetime.now())[5:7]+str(datetime.datetime.now())[8:10]) or last_date == today_std :
                        continue
                    else:
                        fail_list.append(stock)
                        print(stock+"沒更新到最新")
    
        if fail_list == [] or fail_list == fail_stock:
            run_again = 9
            pass
        elif fail_1 == [] :
            fail_1 = fail_list
            continue
        elif fail_2 == [] :
            fail_2 = fail_list
            n2 = 1

        if n2 == 1:
            if len(fail_1) == len(fail_2):
                for i in range(len(fail_1)):
                    if fail_1[i] == fail_2[i]:
                        run_again = 9
                        continue
                    else:
                        run_again = 6
                        fail_1,fail_2 = [],[]
                        n2 = 0
                        break
            else:                
                run_again = 6
                fail_1,fail_2 = [],[]
                n2 = 0
                continue
            
##############################
# 主要運行程式碼-外陸資持有股數 #
#############################   
    fail_1 = fail_2 = [] 
    while run_again == 9:
        column_title =['日期',"證券代號","證券名稱","國際證券編碼",'發行股數','外資及陸資尚可投資股數','全體外資及陸資持有股數','外資及陸資尚可投資比率','全體外資及陸資持股比率','外資及陸資共用法令投資上限比率','陸資法令投資上限比率','與前日異動原因(註)','最近一次上市公司申報外資及陸資持股異動日期']
        column_title_new =["日期",'發行股數','外資及陸資尚可投資股數','全體外資及陸資持有股數','外資及陸資尚可投資比率','全體外資及陸資持股比率','外資及陸資共用法令投資上限比率','陸資法令投資上限比率','與前日異動原因(註)','最近一次上市公司申報外資及陸資持股異動日期']
        switch_ = 0
    
        for i in range(num_stock):
            stock_list = table.row_values(i)[2:]
            number_list = sheet.row_values(i)[2:]
            for stock in stock_list:
                data_list = []
                if stock != '' :
                    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-外資及陸資持股.xls")
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                    xls_file_o = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-比率.xls")
                    xls_df_o = xls_file_o.parse(xls_file_o.sheet_names[0], index_col=[0]) 
    
                    if GetDateList(DateGap(xls_df),xls_df) == 0:
                        continue               # 決定資料天數  
                    time_period = GetDateList(DateGap(xls_df),xls_df) 
    
                    try:
                        for date in time_period : 
                            if date not in np.array(xls_df_o["日期"]):
                                continue
                            pass_ = 0
                            today = datetime.datetime(int(str(date)[:4]),int(str(date)[4:6]),int(str(date)[6:]))
                            if today.weekday() == 6 or today.weekday() == 5 or date in skip_date:
                                continue
    
                            url = "https://www.twse.com.tw/rwd/zh/fund/MI_QFIIS?response=json&date="+str(date)+"&selectType=ALLBUT0999&_=1562078587694"
                            sourcecode = requests.get(url, headers = headers)
                            rawdata = sourcecode.text
                            time.sleep(random.randint(5,8))
                            if "200" in str(sourcecode):
                                switch = 2
                            else:
                                switch = 0
                            if "data" in rawdata:
                                data_want = json.loads(rawdata)["data"]
                                pass_ = 1
                                for data in data_want:
                                    if data[0].strip() == str(int(number_list[stock_list.index(stock)])): 
                                        data.insert(0,str(date))
                                        
                                        for k in [4,5,6,7,8,9,10]:                       
                                            try:
                                                data[k] = float(data[k])
                                            except:
                                                data[k] = int((data[k]).replace(",",""))                                        
                                        
                                        data_list.append(data)
                                        switch = 1
                                        break
                                if switch == 2 and pass_ == 1 :
                                    if date in np.array(xls_df_o["日期"]):
                                        data_list.append([date,0,0,0,0,0,0,0,0,0,0,0,0])                                
                            if switch == 0 : 
                                if date in np.array(xls_df_o["日期"]):
                                    data_list.append([date,0,0,0,0,0,0,0,0,0,0,0,0])                              
    
                        if len(data_list) != 0 :
                            info = pd.DataFrame(data_list,columns = column_title )
                            del info["證券代號"]
                            del info['證券名稱']
                            del info["國際證券編碼"]
                            info.columns = column_title_new
                            below = xls_df.loc[:]
                            newdata = info.append(below,ignore_index=True)
                            newdata.to_excel("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-外資及陸資持股.xls")  # 匯出excel
                            newdata = 0  #防錯機制，假設20180407有資料，因為沒抓到而改用20180408，存檔會變成上一個檔案的資料
                            print(stock+"外資及陸資持股已經下載")
                            print(str(datetime.datetime.now()))
    
                        if switch_ == 0 :
                            today_std = int(info["日期"][0])
                            xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
                            xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                            fail_stock = np.array(xls_df["失敗清單"]).tolist()
                            fail_stock = [x for x in fail_stock if str(x)!= 'nan']
                            switch_ = 1
                    
                    except:
                        print(stock+"外資及陸資持股下載失敗，須注意")
                        print(str(datetime.datetime.now()))
                        continue
               
    
    ###檢查是否有更新到今日日期                    
        n2 = 0
        fail_list = []    
        for j in range(num_stock):
            stock_list = table.row_values(j)[2:]
            number_list = sheet.row_values(j)[2:]
            for stock in stock_list:            
                if stock != '':
                    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(j)[1]+"\\"+stock+"\\"+stock+"-外資及陸資持股.xls")
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                    last_date = int(xls_df["日期"][0])
                    if last_date == int(str(datetime.datetime.now())[:4]+str(datetime.datetime.now())[5:7]+str(datetime.datetime.now())[8:10]) or last_date == today_std :
                        continue
                    else:
                        fail_list.append(stock)
                        print(stock+"沒更新到最新")
    
        if fail_list == [] or fail_list == fail_stock:
            run_again = 10
            pass
        elif fail_1 == [] :
            fail_1 = fail_list
            continue
        elif fail_2 == [] :
            fail_2 = fail_list
            n2 = 1

        if n2 == 1:
            if len(fail_1) == len(fail_2):
                for i in range(len(fail_1)):
                    if fail_1[i] == fail_2[i]:
                        run_again = 10
                        continue
                    else:
                        run_again = 9
                        fail_1,fail_2 = [],[]
                        n2 = 0
                        break
            else:                
                run_again = 9
                fail_1,fail_2 = [],[]
                n2 = 0
                continue

        print("都有更新到最新")
        
#############################
# 主要運行程式碼-融資融券 #
#############################
    fail_1 = fail_2 = [] 
    while run_again == 10:
    
        column_title =["日期","股票代號","股票名稱","融資買進","融資賣出","融資現金償還","融資前日餘額","融資今日餘額","融資限額","融券買進","融券賣出","融券現金償還","融券前日餘額","融券今日餘額","融券限額","資券互抵","註記"]
        column_title_new =["日期","融資買進","融資賣出","融資現金償還","融資前日餘額","融資今日餘額","融資限額","融券買進","融券賣出","融券現金償還","融券前日餘額","融券今日餘額","融券限額","資券互抵","註記"]
        switch_ = 0
        for i in range(num_stock):
            stock_list = table.row_values(i)[2:]
            number_list = sheet.row_values(i)[2:]
            for stock in stock_list:
                data_list = []
                if stock != '':
                    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-融資融券彙總.xls")
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                    xls_file_o = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-比率.xls")
                    xls_df_o = xls_file_o.parse(xls_file_o.sheet_names[0], index_col=[0]) 
    
                    if GetDateList(DateGap(xls_df),xls_df) == 0:
                        continue               # 決定資料天數  
                    time_period = GetDateList(DateGap(xls_df),xls_df) 
    
                    try:
                        for date in time_period :
                            if date not in np.array(xls_df_o["日期"]):
                                continue                            
                            pass_ = 0
                            today = datetime.datetime(int(str(date)[:4]),int(str(date)[4:6]),int(str(date)[6:]))
                            if today.weekday() == 6 or today.weekday() == 5 or date in skip_date:
                                continue
    
                            url = "https://www.twse.com.tw/rwd/zh/marginTrading/MI_MARGN?response=json&date="+str(date)+"&selectType=ALL&_=1562078464225"
                            sourcecode = requests.get(url, headers = headers)
                            rawdata = sourcecode.text
                            time.sleep(random.randint(5,8))
                            if "200" in str(sourcecode):
                                switch = 2
                            else:
                                switch = 0                        
                            if "data" in rawdata:
                                data_want = json.loads(rawdata)["tables"][1]["data"]
                                pass_ = 1
                                for data in data_want:
                                    if data[0].strip() == str(int(number_list[stock_list.index(stock)])): 
                                        data.insert(0,int(date))
                                        
                                        for k in [3,4,5,6,7,8,9,10,11,12,13,14,15]:
                                            try:
                                                data[k] = float(data[k])
                                            except:
                                                data[k] = int((data[k]).replace(",",""))                                        
                                        
                                        data_list.append(data)
                                        switch = 1
                                        break
                                    
                                if switch == 2 and pass_ == 1 :
                                    if date in np.array(xls_df_o["日期"]):
                                        data_list.append([date,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0])                                
                            if switch == 0 :
                                if date in np.array(xls_df_o["日期"]):
                                    data_list.append([date,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]) 
    
                        if len(data_list) != 0 :
                            info = pd.DataFrame(data_list,columns = column_title )
                            del info["股票代號"]
                            del info['股票名稱']
                            below = xls_df.loc[:]
                            below.columns = column_title_new
                            newdata = info.append(below,ignore_index=True)
                            newdata.to_excel("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-融資融券彙總.xls")  # 匯出excel
                            newdata = 0  #防錯機制，假設20180407有資料，因為沒抓到而改用20180408，存檔會變成上一個檔案的資料
                            print(stock+"融資融券已經下載")
                            print(str(datetime.datetime.now()))
    
                        if switch_ == 0 :
                            today_std = int(info["日期"][0])
                            xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
                            xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                            fail_stock = np.array(xls_df["失敗清單"]).tolist()
                            fail_stock = [x for x in fail_stock if str(x)!= 'nan']
                            switch_ = 1
                
                    except:
                        print(stock+"融資融券下載失敗，須注意")
                        print(str(datetime.datetime.now()))
                        continue
               
    
    ###檢查是否有更新到今日日期                    
        n2 = 0
        fail_list = []    
        for j in range(num_stock):
            stock_list = table.row_values(j)[2:]
            number_list = sheet.row_values(j)[2:]
            for stock in stock_list:            
                if stock != '' :
                    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(j)[1]+"\\"+stock+"\\"+stock+"-融資融券彙總.xls")
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                    try:   ###防止沒有資料
                        last_date = int(xls_df["日期"][0])
                    except:
                        continue
                    if last_date == int(str(datetime.datetime.now())[:4]+str(datetime.datetime.now())[5:7]+str(datetime.datetime.now())[8:10]) or last_date == today_std:
                        continue
                    else:
                        fail_list.append(stock)
                        print(stock+"沒更新到最新")
    
        if fail_list == [] or fail_list == fail_stock:
            run_again = 11
            pass
        elif fail_1 == [] :
            fail_1 = fail_list
            continue
        elif fail_2 == [] :
            fail_2 = fail_list
            n2 = 1

        if n2 == 1:
            if len(fail_1) == len(fail_2):
                for i in range(len(fail_1)):
                    if fail_1[i] == fail_2[i]:
                        run_again = 11
                        continue
                    else:
                        run_again = 10
                        fail_1,fail_2 = [],[]
                        n2 = 0
                        break
            else:                
                run_again = 10
                fail_1,fail_2 = [],[]
                n2 = 0
                continue
    
    ##############################
    # 主要運行程式碼-融券借券 #
    #############################
    fail_1 = fail_2 = [] 
    while run_again == 11:
    
        column_title =["日期","股票代號","股票名稱","融券前日餘額","融券賣出","融券買進","融券現券","融券今日餘額","融券限額","借券前日餘額","借券當日賣出","借券當日還券","借券當日調整","借券當日餘額","借券今日可限額","備註"]
        column_title_new =["日期","融券前日餘額","融券賣出","融券買進","融券現券","融券今日餘額","融券限額","借券前日餘額","借券當日賣出","借券當日還券","借券當日調整","借券當日餘額","借券今日可限額","備註"]
        switch_= 0
        for i in range(num_stock):
            stock_list = table.row_values(i)[2:]
            number_list = sheet.row_values(i)[2:]
            for stock in stock_list:
                data_list = []
                if stock != '' :# and stock not in ["日月光投控","群益期","千興","昇陽光電","昱晶","中時","東訊","倫飛","互億"]:
                    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-融券借券彙總.xls")
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                    xls_file_o = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-比率.xls")
                    xls_df_o = xls_file_o.parse(xls_file_o.sheet_names[0], index_col=[0])                     
                                     
                    if GetDateList(DateGap(xls_df),xls_df) == 0:
                        continue               # 決定資料天數  
                    time_period = GetDateList(DateGap(xls_df),xls_df)  

                    try:
                        for date in time_period :
                            if date not in np.array(xls_df_o["日期"]):
                                continue                            
                            pass_ = 0
                            today = datetime.datetime(int(str(date)[:4]),int(str(date)[4:6]),int(str(date)[6:]))
                            if today.weekday() == 6 or today.weekday() == 5 or date in skip_date:
                                continue
                            url = "http://www.twse.com.tw/rwd/zh/marginTrading/TWT93U?response=json&date="+str(date)+"&_=1497966716847"
                            sourcecode = requests.get(url, headers = headers)
                            rawdata = sourcecode.text
                            time.sleep(random.randint(5,8))
                            if "200" in str(sourcecode):
                                switch = 2
                            else:
                                switch = 0
                            if "data" in rawdata:
                                data_want = json.loads(rawdata)["data"]
                                pass_ = 1
                                for data in data_want:
                                    if data[0].strip() == str(int(number_list[stock_list.index(stock)])): 
                                        data.insert(0,str(date))
                                        
                                        for k in [3,4,5,6,7,8,9,10,11,12,13,14]:
                                            try:
                                                data[k] = float(data[k])
                                            except:
                                                data[k] = int((data[k]).replace(",",""))                                        
                        
                                        data_list.append(data)
                                        switch = 1
                                        break
                                if switch == 2 and pass_ == 1 :
                                    if date in np.array(xls_df_o["日期"]):
                                        data_list.append([date,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0])                                
                            if switch == 0 :
                                if date in np.array(xls_df_o["日期"]):
                                    data_list.append([date,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]) 
    
                        if len(data_list) != 0 :
                            info = pd.DataFrame(data_list,columns = column_title )
                            del info["股票代號"]
                            del info['股票名稱']
                            below = xls_df.loc[:]
                            below.columns = column_title_new
                            newdata = info.append(below,ignore_index=True)
                            newdata.to_excel("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-融券借券彙總.xls")  # 匯出excel
                            newdata = 0  #防錯機制，假設20180407有資料，因為沒抓到而改用20180408，存檔會變成上一個檔案的資料
                            print(stock+"-融券借券已經下載")
                            print(str(datetime.datetime.now()))
    
                        if switch_ == 0 :
                            today_std = int(info["日期"][0])
                            xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
                            xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                            fail_stock = np.array(xls_df["失敗清單"]).tolist()
                            fail_stock = [x for x in fail_stock if str(x)!= 'nan']
                            switch_ = 1
                  
                    except:
                        print(stock+"融券借券下載失敗，須注意")
                        print(str(datetime.datetime.now()))                    
                        continue
               
    
    ###檢查是否有更新到今日日期                     
        n2 = 0
        fail_list = []    
        for j in range(num_stock):
            stock_list = table.row_values(j)[2:]
            number_list = sheet.row_values(j)[2:]
            for stock in stock_list:            
                if stock != '' :
                    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(j)[1]+"\\"+stock+"\\"+stock+"-融券借券彙總.xls")
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0]) 
                    try:   ###防止沒有資料
                        last_date = int(xls_df["日期"][0])
                    except:
                        continue
                    if last_date == int(str(datetime.datetime.now())[:4]+str(datetime.datetime.now())[5:7]+str(datetime.datetime.now())[8:10]) or last_date == today_std:
                        continue
                    else:
                        fail_list.append(stock)
                        print(stock+"沒更新到最新")
    
        if fail_list == [] or fail_list == fail_stock:
            run_again,enter = 7,0
            pass
        elif fail_1 == [] :
            fail_1 = fail_list
            continue
        elif fail_2 == [] :
            fail_2 = fail_list
            n2 = 1

        if n2 == 1:
            if len(fail_1) == len(fail_2):
                for i in range(len(fail_1)):
                    if fail_1[i] == fail_2[i]:
                        run_again,enter = 7,0
                        continue
                    else:
                        run_again = 11
                        fail_1,fail_2 = [],[]
                        n2 = 0
                        break
            else:                
                run_again = 11
                fail_1,fail_2 = [],[]
                n2 = 0
                continue
    
##############################
# 主要運行程式碼-檢查資訊正確 #
#############################      
    while run_again == 7: 
        if datetime.datetime.now().weekday() in [0,1,2,3,4] and enter == 0: 
            if int(str(datetime.datetime.now())[11:13]+str(datetime.datetime.now())[14:16]) >= 1435:
                enter = enter + 1
                run_again = 1
                break            

        else:
            time.sleep(300)
