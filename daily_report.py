# -*- coding: utf-8 -*-
"""
Created on Tue Feb  5 10:15:51 2019

@author: YAO
"""

import xlrd
import numpy as np
import pandas as pd
from datetime import datetime
import price_slope_daily_new
import math
import more_itertools as mit


def excel_locate():
    for i in range(len(xls_df_basic)):
        if xls_df_basic["股票名稱"][i] == stock:
            break
    return i

def erase_for_big(xls_file_big_d,xls_df_report):
    
    xls_df_date = set(list(xls_df_report["日期"]))
    xls_df_big_date = set(list(xls_file_big_d["日期"]))
    diff = xls_df_big_date.difference(xls_df_date)
    diff = list(diff)   
    diff.sort(reverse = False)
    
    for i in range(xls_file_big_d.index[0],xls_file_big_d.index[-1]+1):
        if xls_file_big_d["日期"][i] in diff :
            xls_file_big_d.drop([i],axis = 0, inplace = True)
    xls_file_big_d = xls_file_big_d.sort_values(by = "日期",ascending = True)
    xls_file_big_d.index = range(xls_df_report.index[-1]-len(xls_file_big_d)+1,xls_df_report.index[-1]+1)
    
    return xls_file_big_d

    
xls_file_cate = pd.ExcelFile("D:\\Stock Investment\\個股分類表.xlsx")
xls_df_cate = xls_file_cate.parse(xls_file_cate.sheet_names[0], index_col=[0])
    
file_industry = xlrd.open_workbook("D:\\Stock Investment\\股票分類-下載用(去除重複)-分類.xlsx")
table_industry = file_industry.sheets()[0]

xls_file = pd.ExcelFile("D:\\Stock Investment\\產業大類-名稱.xlsx")
table_category = xls_file.parse(xls_file.sheet_names[0], index_col=[0])


i =  int(input("1-上市/2-上櫃"))

if i == 1 :
    file = xlrd.open_workbook(r"D:\Stock Investment\股票分類-下載用(去除重複)-上市.xlsx")
    table = file.sheets()[0]
    sheet = file.sheets()[1]
    num_stock = 32
    switch_big = 0
    
elif i == 2:
    file = xlrd.open_workbook(r"D:\Stock Investment\股票分類-下載用(去除重複)-上櫃.xlsx")
    table = file.sheets()[0]
    sheet = file.sheets()[1] 
    num_stock = 30
    switch_big = 1
    
xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤-每日清單.xlsx")
xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])                
date_all = np.array(xls_file_big_d["日期"]).tolist()
date_all.sort()
target_date = 20230101
delete_date = date_all[date_all.index(target_date)-15]
yesterday_std = date_all[date_all.index(target_date)-1]

for date in date_all:
    if date < target_date or date >= 20230710:
        continue
    switch = 1

    if switch_big == 0 :                    
        xls_file_basic = pd.ExcelFile("D:\\Stock Investment\\技術分析表-上市.xlsx")
        xls_df_basic = xls_file_basic.parse(xls_file_basic.sheet_names[0], index_col=[0])                  
    elif switch_big == 1 :    
        xls_file_basic = pd.ExcelFile("D:\\Stock Investment\\技術分析表-上櫃.xlsx")
        xls_df_basic = xls_file_basic.parse(xls_file_basic.sheet_names[0], index_col=[0]) 
                        
    for i in range(num_stock):
        stock_list = table.row_values(i)[2:]
        number_list = sheet.row_values(i)[2:]
        category = table.row_values(i)[1]
        for stock in stock_list:                           
            if stock != "" : 
                if switch_big == 0 :            
                    if "工業" in category:
                        loca = category.strip().index("工")
                        category = category.strip()[:loca]
                    elif "業" in category and category != "觀光事業" and category != "航運業" :
                        loca = category.strip().index("業")
                        category = category.strip()[:loca] 

                elif switch_big == 1 :   
                    if "工業" in category:
                        loca = category.strip().index("工")
                        category = category.strip()[:loca]
                    elif "業" in category and category != "觀光事業" and category != "農業科技業" :
                        loca = category.strip().index("業")
                        category = category.strip()[:loca]
                    elif  category == "農業科技業" :
                        category = category.strip()[:4] 
                    if category == "金融保險" :
                        category = "金融"
                    elif category == "航運業" :
                        category = "航運"
                
                try:
                    xls_df_day = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)
                    xls_df_week = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-week.csv",engine='python', encoding = "big5",index_col=0)
                    date_basic = xls_df_day["日期"][0]
                except:
                    continue
                for k in range(len(xls_df_day)):
                    if date in np.array(xls_df_day["日期"]):
                       if int(xls_df_day["日期"][k]) == date:
                            xls_df_day = xls_df_day[:k+1]
                            break
                    else:
                        continue
          
                if switch_big == 0 :                    
                    xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\資金流向-市-每日清單.xlsx")
                    xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])    
                    xls_df_big = erase_for_big(xls_file_big_d,xls_df_day)                   
                elif switch_big == 1 :     
                    xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\資金流向-櫃-每日清單.xlsx")
                    xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])    
                    xls_df_big = erase_for_big(xls_file_big_d,xls_df_day) 
            
                locate = excel_locate()
                days_week,days_day,switch = 50,60,1

                price_slope_daily_new.price_range_day(stock,xls_df_day,xls_df_basic,locate)
                price_slope_daily_new.perb_data_show(stock,xls_df_day,xls_df_basic,locate)
                price_slope_daily_new.meet_20days_standard_within_60days(stock,xls_df_day,xls_df_big,xls_df_basic,locate)
                price_slope_daily_new.meet_20days_standard_today(stock,xls_df_day,xls_df_big,xls_df_basic,locate)
                price_slope_daily_new.up_down_day(stock,xls_df_day,xls_df_basic,locate)
                price_slope_daily_new.volume_day(stock,xls_df_day,xls_df_basic,locate)
                price_slope_daily_new.signal_wait_for_5ma(stock,xls_df_day,xls_df_basic,locate)                                        
                price_slope_daily_new.industry_stock(stock,xls_df_day,xls_df_basic,xls_df_cate,table_industry,table_category,locate,category,xls_df_big)

                xls_df_basic["收盤價"][locate] = xls_df_day["日收盤價"][len(xls_df_day["日收盤價"])-1]          
                switch = 1
                print(stock+str(date)+"完成")

                
    if switch == 1 and switch_big == 0:     
        xls_df_basic.to_excel(r"C:\Users\User\Desktop\每日計算報告\上市總表"+str(date)+"1.xls")    
        print(str(date)+"總表已經下載")
        print(str(datetime.now()))
                
    if switch == 1 and switch_big == 1:         
        xls_df_basic.to_excel(r"C:\Users\User\Desktop\每日計算報告\上櫃總表"+str(date)+"1.xls")    
        print(str(date)+"總表已經下載")
        print(str(datetime.now()))


    columns = ["標題一","標題二","標題三","標題四","標題四十一","標題五","標題六","標題七","標題八","標題九","標題十","標題十一","標題十二","標題十三","標題十四","標題十五","標題十六","標題十七","標題十八","標題十九","標題二十","標題二十一","標題二十二","標題二十三","標題二十四","標題二十五","標題二十六","標題二十七","標題二十八","標題二十九","標題三十","標題三十一","標題三十二","標題三十三","標題三十四","標題三十五","標題三十六","標題三十七","標題三十八","因子十六","標題四十"
            	,"標題十一-長期","標題十一-短期","標題十一-長期排名","標題十一-短期排名","標題十二-長期","標題十二-短期","標題十二-長期排名","標題十二-短期排名","標題十三-長期","標題十三-短期","標題十三-長期排名","標題十三-短期排名","標題十四-長期","標題十四-短期","標題十四-長期排名","標題十四-短期排名","標題十五-長期","標題十五-短期","標題十五-長期排名","標題十五-短期排名","標題十六-長期","標題十六-短期","標題十六-長期排名","標題十六-短期排名","標題十七-長期","標題十七-短期","標題十七-長期排名","標題十七-短期排名","標題十九-長期","標題十九-短期"
                ,"標題十九-長期排名","標題十九-短期排名","標題十八-長期","標題十八-短期","標題十八-長期排名","標題十八-短期排名","標題二十-長期","標題二十-短期","標題二十-長期排名","標題二十-短期排名","標題二十一-長期","標題二十一-短期","標題二十一-長期排名","標題二十一-短期排名","標題二十二-長期","標題二十二-短期","標題二十二-長期排名","標題二十二-短期排名","標題二十三-長期","標題二十三-短期","標題二十三-長期排名","標題二十三-短期排名","標題二十四-長期","標題二十四-短期"
                ,"標題二十四-長期排名","標題二十四-短期排名","標題二十五-長期","標題二十五-短期","標題二十五-長期排名","標題二十五-短期排名","標題二十六-長期","標題二十六-短期","標題二十六-長期排名","標題二十六-短期排名","標題二十八-長期","標題二十八-短期","標題二十八-長期排名","標題二十八-短期排名","標題二十九-長期","標題二十九-短期","標題二十九-長期排名","標題二十九-短期排名","標題三-長期","標題三-短期","標題三-長期排名","標題三-短期排名","標題四-長期","標題四-短期","標題四-長期排名","標題四-短期排名","標題四十一-長期","標題四十一-短期","標題四十一-長期排名"
                ,"標題四十一-短期排名","標題七-長期","標題七-短期","標題七-長期排名","標題七-短期排名","標題八-長期","標題八-短期","標題八-長期排名","標題八-短期排名","標題三十二-長期","標題三十二-短期","標題三十二-長期排名","標題三十二-短期排名","標題三十三-長期","標題三十三-短期","標題三十三-長期排名","標題三十三-短期排名","標題三十四-長期","標題三十四-短期","標題三十四-長期排名","標題三十四-短期排名","標題四十二",'標題四十三',"標題四十四","標題四十五","標題四十六","標題四十七","標題四十八",'標題四十九',"標題五十","標題五十一","標題五十二","標題五十三","標題五十四","標題五十五","標題五十六","標題五十七","標題五十八","標題五十九","標題六十"
                ,"標題六十一","標題六十二",'標題六十三',"標題六十四","標題六十五","標題六十六","標題六十七",'標題六十八',"標題六十九"] 

    for weight in ["股數"]:              
        category_stock = {}
        for i in range(len(xls_df_basic)):
            try:
                if type(xls_df_basic[weight+"股"][i]) == str:
                    for key in eval(xls_df_basic[weight+"股"][i]):
                        stock_list = []   
                        if key in category_stock.keys():
                            for stock in category_stock[key]:
                                stock_list.append(stock)
                            stock_list.append(xls_df_basic["標題三十七"][i])
                            category_stock.update({key:stock_list})
                        else:
                            category_stock.setdefault(key,[xls_df_basic["標題三十七"][i]])
            except:
                continue
   
        if weight == "股數":
            category_uni_big,category_price_big,category_big_category,category_big_hold,category_big_hold_2 = {},{},{},{},{}
            for i in range(len(xls_df_basic)):
                if type(xls_df_basic["大對應"][i]) == str:
                    for key in eval(xls_df_basic["大對應"][i]):
                        category_big_category.setdefault(key,eval(xls_df_basic["大對應"][i])[key])
                    for key in eval(xls_df_basic["大持有-今日"][i]):
                        if key not in category_big_hold.values():                                                        
                            category_big_hold.setdefault(key,eval(xls_df_basic["大持有-今日"][i])[key])  
                    for key in eval(xls_df_basic["標題二"][i]):
                        if key not in category_big_hold_2.values():                                                        
                            category_big_hold_2.setdefault(key,eval(xls_df_basic["標題二"][i])[key])  
                    for key in eval(xls_df_basic["標題七"][i]):
                        if key not in category_uni_big.values():                                                        
                            category_uni_big.setdefault(key,eval(xls_df_basic["標題七"][i])[key])  
                    for key in eval(xls_df_basic["標題八"][i]):
                        if key not in category_price_big.values():                                                        
                            category_price_big.setdefault(key,eval(xls_df_basic["標題八"][i])[key])  
                         
        total_list = []   
        for key in category_stock:
            stock_list = category_stock[key]
                                                          
            for stock in stock_list:
                data_list,message = [],0
                for i in range(len(xls_df_basic)): 
                    if xls_df_basic["標題三十七"][i] == stock:   
                        try:
                            if math.isnan(xls_df_basic["標題七十"][i]) == True:    
                                continue
                        except:
                            pass
                        if weight == "產業":
                            for item in ["標題三","標題四","標題五","標題六","標題七十","標題七","標題八"]:
                                xls_df_basic[item][i] = xls_df_basic[item][i].replace("nan","0")
                            try:
                                data_list.append(category_big_hold[key])
                                data_list.append(category_big_hold_2[key])
                            except:
                                data_list.append("")
                                data_list.append("")
                            data_list.append(eval(xls_df_basic["標題三"][i])[key])
                            data_list.append(eval(xls_df_basic["標題四"][i])[key])
                            data_list.append(eval(xls_df_basic["標題七十"][i])[key])
                            data_list.append(eval(xls_df_basic["標題五"][i])[key])
                            data_list.append(eval(xls_df_basic["標題六"][i])[key])
                            data_list.append(eval(xls_df_basic["標題七"][i])[key])
                            data_list.append(eval(xls_df_basic["標題八"][i])[key])
                            try:
                                data_list.append(category_big_category[key]) 
                            except:
                                data_list.append("")
                        else:
                            data_list.append("")
                            data_list.append("")
                            data_list.append("")
                            data_list.append("")
                            data_list.append("") 
                            data_list.append("")
                            data_list.append("") 

                        data_list.append(key)                        

                        for item in [weight+"因子十七",weight+"因子十六",weight+"因子一",weight+"因子九",weight+"因子二",weight+"因子三",weight+"因子四",weight+"因子五",weight+"因子六",weight+"因子七",weight+"因子八",weight+"因子十九",weight+"因子十八",weight+"因子十四",weight+"因子十三",weight+"因子十二",weight+"因子十",weight+"因子十一", weight+"因子二十二",weight+"因子二十一",weight+"因子二十",weight+"因子二十二三", weight+"因子十五", weight+"因子十四"]:
                            xls_df_basic[item][i] = xls_df_basic[item][i].replace("nan","0")
                        data_list.append(eval(xls_df_basic[weight+"因子一"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子九"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子二"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子三"][i])[key])                        
                        data_list.append(eval(xls_df_basic[weight+"因子一"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子十"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子十八"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子十二"][i])[key])  
                        data_list.append(eval(xls_df_basic[weight+"因子十四"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子十三"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子十一"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子四"][i])[key]) 
                        data_list.append(eval(xls_df_basic[weight+"因子八"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子五"][i])[key]) 
                        data_list.append(eval(xls_df_basic[weight+"因子十九"][i])[key])      
                        data_list.append(eval(xls_df_basic[weight+"因子六"][i])[key])  
                        data_list.append(eval(xls_df_basic[weight+"因子七"][i])[key]) 
                        data_list.append(eval(xls_df_basic[weight+"因子十七"][i])[key])  
                        data_list.append(eval(xls_df_basic[weight+"因子十六"][i])[key])
                        try:
                            data_list.append(eval(xls_df_basic[weight+"因子二十二"][i])[key])
                            data_list.append(eval(xls_df_basic[weight+"因子二十一"][i])[key]) 
                        except:
                            data_list.append([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
                            data_list.append([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]) 
                        
                        data_list.append(eval(xls_df_basic[weight+"因子二十"][i])[key])
                        data_list.append(eval(xls_df_basic[weight+"因子二十二三"][i])[key])  
                        data_list.append([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])  
#                        data_list.append(eval(xls_df_basic[weight+"因子十五"][i])[key])  
                        data_list.append(xls_df_basic["漲跌幅"][i])                
                        data_list.append(xls_df_basic["3日漲跌幅"][i])
                        data_list.append(xls_df_basic["標題三十七"][i])

                        if type(xls_df_basic["訊號發動盤堅天數"][i]) == str:
                            message = "訊號發動"                    
                        if xls_df_basic["訊號-今日"][i] == "V":
                            if message != 0 :
                                message = message + "," + "訊號-今日"
                            else:
                                message = "訊號-今日"
                        if xls_df_basic["訊號-10日內"][i] == "V":
                            if message != 0 :
                                message = message + "," + "訊號-10日內"
                            else:
                                message = "訊號-10日內"                
                        data_list.append(message)  
                        print(message)
                        if message != 0 and "發動" in message:
                            data_list.append("訊:"+str(xls_df_basic["訊號價"][i])+" / 收:"+str(xls_df_basic["收盤價"][i]))  
                            data_list.append(xls_df_basic["訊號發動盤堅天數"][i]) 
                        else:
                            data_list.append(xls_df_basic["收盤價"][i])  
                            data_list.append("") 
                        for j in range(132):
                            data_list.append("")
                        total_list.append(data_list)
     
        data_category = pd.DataFrame(total_list ,columns = columns)      
    
    stock_list,switch = [],0
    for i in range(len(data_category)):
        if data_category["標題一"][i] == "" and data_category["標題三十七"][i] in stock_list:
            data_category.drop([i],axis = 0, inplace = True)
            switch = 1
    if switch == 1:
        data_category.index = range(len(data_category))        

    for k in range(len(data_category)):
        for item in ["標題十一","標題十二","標題十三","標題十四","標題二十二","標題二十三","標題十九","標題四十一","標題十五","標題十六","標題十七","標題十八","標題二十四","標題二十五","標題二十六","標題二十","標題二十一","標題三","標題四","標題三十二","標題三十三","標題三十四","標題二十九","標題二十八","標題八","標題七"]:            
            hundred_amount = 0
            try:
                original = eval(data_category[item][k])
            except:
                original = data_category[item][k]
            
            if original == "":
                data_category[item+"-長期"][k] = 0
                data_category[item+"-短期"][k] = 0 
                continue
            
            switch,time = 0,0
            for factor in original:
                if factor == 100 :
                    continue
                elif factor >= 90 and factor != original[-1]:
                    time = time + 1
                elif factor >= 90 and factor == original[-1]:
                    time = time + 1
                    switch = 1
            
            if switch == 1 and time == 1 :
                original[-1] = 100

            for num in range(len(original)):
                if original[num] >= 95 :
                    original[num] = 100

            if len([i for i,v in enumerate(original) if 100 >v>= 90]+[i for i,v in enumerate(original) if v== 100]) == len(original):
                for num in range(len(original)):
                    original[num] = 100  
            else:
                hundred_list = [list(group) for group in mit.consecutive_groups([i for i,v in enumerate(original) if v== 100])]
                if len(hundred_list) > 1 :
                    for p in range(len(hundred_list)-1):
                        time_90 = 0
                        if p == 0 and hundred_list[p][0] != 0:
                            head = 0
                            tail = hundred_list[p][0]
                        else:   
                            head = hundred_list[p][-1]
                            tail = hundred_list[p+1][0]

                        for num in range(head+1,tail):
                            if original[num] >= 90:
                                time_90 = time_90 + 1   
                        if time_90 == len(range(head+1,tail)):
                            for num in range(head+1,tail):
                                original[num] = 100
                                
                elif len(hundred_list) == 1 and hundred_list[0][0] != 0:
                        head,time_90 = 0,0
                        tail = hundred_list[0][0]     
                        time_90 = 0
                        for num in range(head,tail):
                            if original[num] >= 90:
                                time_90 = time_90 + 1   
                        if time_90 == len(range(head,tail)):
                            for num in range(head,tail):
                                original[num] = 100
                                                  
##容忍90以下兩次！ 
            up_list = original[:]
            if len(up_list) == 1 or up_list[0] == 100 and up_list[1] == 100 :
                data_category[item+"-短期"][k] = 0 
                data_category[item+"-短期排名"][k] = 999	
                
            else:
                list_dis = []
                for h in range(len(up_list)-1):
                    list_dis.append(round(up_list[h]-up_list[h+1],3))
                
                data_category[item+"-短期"][k] = list_dis
                
            hundred_amount = 0 
            for num in range(len(original)):
                switch = 1
                for p in range(num):
                    if original[p] < 90 : 
                        switch = 0
                        break
                if switch == 1 and original[num] == 100 :
                    hundred_amount = hundred_amount + 1
                    
            data_category[item+"-長期"][k] = hundred_amount
       
    switch_category = 0
    for k in range(len(data_category)):
        if data_category["標題九"][k] != "":
            switch_category = 1

    if switch_category == 1:            
        end = ""
        for k in range(len(data_category)):
            if data_category["標題九"][k] == "":
                end = k
                break
        if end == "":
            end = len(data_category)
        
        for item in ["標題十一","標題十二","標題十三","標題十四","標題二十二","標題二十三","標題十五","標題十六","標題十七","標題十八","標題十九","標題四十一","標題二十四","標題二十五","標題二十六","標題二十","標題二十一","標題三","標題四","標題三十二","標題三十三","標題三十四","標題二十九","標題二十八","標題八","標題七"]:                  
            data_list = list(set(data_category[item+"-長期"][:end].tolist()))
            data_list.sort(reverse = True)                
            
            for j in range(len(data_list)):
                for k in range(end):
                    if data_category[item+"-長期"][k] == data_list[j] :
                        data_category[item+"-長期排名"][k] = j + 1	
            
            big_category = {}
            if item in ["標題三十二"]:
                for key in set(data_category["標題十"][:end].tolist()):
                    big_category[key] = 9999
                for key in big_category.keys():
                    for k in range(end):
                        if data_category["標題十"][k] == key and data_category[item+"-長期排名"][k] < big_category[key]:
                            big_category[key] = data_category[item+"-長期排名"][k]
                for k in range(end):
                    data_category[item+"-長期排名"][k] = big_category[data_category["標題十"][k]]

            elif item in ["標題三十三"]:
                for key in set(data_category["標題三十七"][:end].tolist()):
                    big_category[key] = 9999
                for key in big_category.keys():
                    for k in range(end):
                        if data_category["標題三十七"][k] == key and data_category[item+"-長期排名"][k] < big_category[key]:
                            big_category[key] = data_category[item+"-長期排名"][k]
                for k in range(end):
                    data_category[item+"-長期排名"][k] = big_category[data_category["標題三十七"][k]]
                    
            elif item in ["標題三","標題四","大增減率","標題八","標題七"]:
                for key in set(data_category["標題九"][:end].tolist()):
                    big_category[key] = 9999
                for key in big_category.keys():
                    for k in range(end):
                        if data_category["標題九"][k] == key and data_category[item+"-長期排名"][k] < big_category[key]:
                            big_category[key] = data_category[item+"-長期排名"][k]
                for k in range(end):
                    data_category[item+"-長期排名"][k] = big_category[data_category["標題九"][k]]
                                                
        for item in ["標題十一","標題十二","標題十三","標題十四","標題二十二","標題二十三","標題十五","標題十六","標題十七","標題十八","標題十九","標題四十一","標題二十四","標題二十五","標題二十六","標題二十","標題二十一","標題三","標題四","標題三十二","標題三十三","標題三十四","標題二十九","標題二十八","標題八","標題七"]:
       	    data_list = data_category[item+"-短期"][:end].tolist()
            while True:
                if 0 in data_list:
                    data_list.remove(0)
                else:
                    break
                    
            data_list.sort(key=lambda x: (-x[0], -x[1], -x[2], -x[3], -x[4], -x[5], -x[6], -x[7], -x[8], -x[9], -x[10], -x[11], -x[12], -x[13], -x[14], -x[15], -x[16], -x[17], -x[18]))
            
            for j in range(len(data_list)):
                for k in range(end):
                    if data_category[item+"-短期"][k] == data_list[j] :
                        data_category[item+"-短期排名"][k] = j + 1	
            
        for k in range(end):                                                                                                                                     
            data_category["標題四十四"][k] =  data_category["標題十三-短期排名"][k] + data_category["標題十三-長期排名"][k] + data_category["標題十四-短期排名"][k] + data_category["標題十四-長期排名"][k]
            data_category["標題四十五"][k] =  data_category["標題十一-短期排名"][k] + data_category["標題十一-長期排名"][k] + data_category["標題十二-短期排名"][k] + data_category["標題十二-長期排名"][k]

            if data_category["標題十五"][k][0] == 100 and data_category["標題十五"][k][1] == 100:
                industry_stock = data_category["標題十五-長期排名"][k]
            else:
                industry_stock = data_category["標題十五-短期排名"][k]
            if data_category["標題十六"][k][0] == 100 and data_category["標題十六"][k][1] == 100:
                industry_money = data_category["標題十六-長期排名"][k]
            else:
                industry_money = data_category["標題十六-短期排名"][k]    
            data_category["標題四十七"][k] =  industry_stock + industry_money

            if data_category["標題十七"][k][0] == 100 and data_category["標題十七"][k][1] == 100:
                industry_stock = data_category["標題十七-長期排名"][k]
            else:
                industry_stock = data_category["標題十七-短期排名"][k]
            if data_category["標題十八"][k][0] == 100 and data_category["標題十八"][k][1] == 100:
                industry_money = data_category["標題十八-長期排名"][k]
            else:
                industry_money = data_category["標題十八-短期排名"][k]    
            if data_category["標題十九"][k][0] == 100 and data_category["標題十九"][k][1] == 100:
                industry_stock_growth = data_category["標題十九-長期排名"][k]
            else:
                industry_stock_growth = data_category["標題十九-短期排名"][k]    
            data_category["標題四十八"][k] =  industry_money + industry_stock
            data_category["標題四十九"][k] =  industry_stock_growth
            
            if data_category["標題二十二"][k][0] == 100 and data_category["標題二十二"][k][1] == 100:
                industry_stock = data_category["標題二十二-長期排名"][k]
            else:
                industry_stock = data_category["標題二十二-短期排名"][k]
            if data_category["標題二十三"][k][0] == 100 and data_category["標題二十三"][k][1] == 100:
                industry_money = data_category["標題二十三-長期排名"][k]
            else:
                industry_money = data_category["標題二十三-短期排名"][k]             
            data_category["標題四十六"][k] =  industry_stock #+ industry_money
            
            if data_category["標題二十"][k][0] == 100 and data_category["標題二十"][k][1] == 100:
                industry_stock = data_category["標題二十-長期排名"][k]
            else:
                industry_stock = data_category["標題二十二-短期排名"][k]
            if data_category["標題二十一"][k][0] == 100 and data_category["標題二十一"][k][1] == 100:
                industry_money = data_category["標題二十一-長期排名"][k]
            else:
                industry_money = data_category["標題二十一-短期排名"][k]              
            data_category["標題五十四"][k] =  industry_stock + industry_money
            
            if data_category["標題二十四"][k][0] == 100 and data_category["標題二十四"][k][1] == 100:
                industry_stock = data_category["標題二十四-長期排名"][k]
            else:
                industry_stock = data_category["標題二十四-短期排名"][k]
            if data_category["標題二十五"][k][0] == 100 and data_category["標題二十五"][k][1] == 100:
                industry_money = data_category["標題二十五-長期排名"][k]
            else:
                industry_money = data_category["標題二十五-短期排名"][k]                
            if data_category["標題二十六"][k][0] == 100 and data_category["標題二十六"][k][1] == 100:
                industry_stock_growth = data_category["標題二十六-長期排名"][k]
            else:
                industry_stock_growth = data_category["標題二十六-短期排名"][k]
            data_category["標題五十"][k] = industry_stock #+ industry_money
            data_category["標題五十一"][k] = industry_stock_growth

            if data_category["標題二十九"][k][0] == 100 and data_category["標題二十九"][k][1] == 100:
                industry_stock = data_category["標題二十九-長期排名"][k]
            else:
                industry_stock = data_category["標題二十九-短期排名"][k]
            if data_category["標題二十八"][k][0] == 100 and data_category["標題二十八"][k][1] == 100:
                industry_money = data_category["標題二十八-長期排名"][k]
            else:
                industry_money = data_category["標題二十八-短期排名"][k]    
            data_category["標題五十三"][k] =  industry_stock + industry_money
            
            if data_category["標題三"][k][0] == 100 and data_category["標題三"][k][1] == 100:
                industry_stock = data_category["標題三-長期排名"][k]
            else:
                industry_stock = data_category["標題三-短期排名"][k]
            if data_category["標題四"][k][0] == 100 and data_category["標題四"][k][1] == 100:
                industry_money = data_category["標題四-長期排名"][k]
            else:
                industry_money = data_category["標題四-短期排名"][k]  
            if data_category["標題四十一"][k][0] == 100 and data_category["標題四十一"][k][1] == 100:
                industry_stock_growth = data_category["標題四十一-長期排名"][k]
            else:
                industry_stock_growth = data_category["標題四十一-短期排名"][k]                
            data_category["標題四十三"][k] =  industry_stock_growth              
            data_category["標題四十二"][k] =  industry_stock + industry_money

            if data_category["標題八"][k][0] == 100 and data_category["標題八"][k][1] == 100:
                industry_stock = data_category["標題八-長期排名"][k]
            else:
                industry_stock = data_category["標題八-短期排名"][k]
            if data_category["標題七"][k][0] == 100 and data_category["標題七"][k][1] == 100:
                industry_money = data_category["標題七-長期排名"][k]
            else:
                industry_money = data_category["標題七-短期排名"][k]    
            data_category["標題五十二"][k] =  industry_stock + industry_money
            
            if data_category["標題三十二"][k][0] == 100 and data_category["標題三十二"][k][1] == 100:
                industry_stock = data_category["標題三十二-長期排名"][k]
            else:
                industry_stock = data_category["標題三十二-短期排名"][k]
            if data_category["標題三十三"][k][0] == 100 and data_category["標題三十三"][k][1] == 100:
                industry_money = data_category["標題三十三-長期排名"][k]
            else:
                industry_money = data_category["標題三十三-短期排名"][k] 
            if data_category["標題三十四"][k][0] == 100 and data_category["標題三十四"][k][1] == 100:
                industry_three = data_category["標題三十四-長期排名"][k]
            else:
                industry_three = data_category["標題三十四-短期排名"][k] 
            data_category["標題五十五"][k] =  industry_stock + industry_money + industry_three
            
        for item in ["因子二十三","因子二十四","因子二十五","因子二十六","因子二十七","因子二十八","因子二十九","因子三十","因子三十一","因子三十二","因子三十三","因子三十四","因子三十五","因子三十六","因子三十七"]:
            rank_list = np.array(data_category[item+"總積分"][:end]).tolist()
            rank_list = list(set(rank_list))
            rank_list.sort()
            for k in range(len(rank_list)):
                for p in range(len(data_category[item+"總排名"][:end])):
                    if data_category[item+"總積分"][p] == rank_list[k]:
                        data_category[item+"總排名"][p] = k + 1
     
        total_data = []
        for k in range(end):            
            total_data.append([data_category["標題三十七"][k],data_category["標題十"][k],data_category["標題六十"][k],data_category["標題六十一"][k],data_category["標題六十七"][k],data_category["標題五十九"][k],data_category["標題六十九"][k],data_category["標題六十六"][k]])            
        
        total_data.sort(key=lambda x: (x[2], x[3], x[4], x[5]))

        total_data_arrange = []
        for j in range(len(total_data)):
            for k in range(end):
                if [data_category["標題三十七"][k],data_category["標題十"][k],data_category["標題六十"][k],data_category["標題六十一"][k],data_category["標題六十七"][k],data_category["標題五十九"][k],data_category["標題六十九"][k],data_category["標題六十六"][k]] == total_data[j]:
                    total_data_arrange.append(np.array(data_category).tolist()[k])
                    break

        data_category = pd.DataFrame(total_data_arrange ,columns = columns)             
        if switch_big == 0 :        
            data_category.to_excel("C:\\Users\\User\Desktop\\每日計算報告\\上市結果報表\\"+str(date)+".xls")    
            xls_df_strong = pd.ExcelFile("C:\\Users\\User\\Desktop\\每日計算報告\\強勢股列表-上市"+str(yesterday_std)+".xls")
            xls_df_strong = xls_df_strong.parse(xls_df_strong.sheet_names[0], index_col=[0]) 
            stock_list = list(set(xls_df_strong["標題三十七"].tolist()))

        elif switch_big == 1 : 
            data_category.to_excel(r"C:\Users\User\Desktop\每日計算報告\上櫃結果報表"+str(date)+".xls")  
            xls_df_strong = pd.ExcelFile(r"C:\Users\User\Desktop\每日計算報告\強勢股列表-上櫃"+str(yesterday_std)+".xls")
            xls_df_strong = xls_df_strong.parse(xls_df_strong.sheet_names[0], index_col=[0])                 
            stock_list = list(set(xls_df_strong["標題三十七"].tolist()))
                
    valid_stock = []
    for k in range(len(data_category)):
        if data_category["標題三十五"][k] >= 6 and data_category["標題三十七"][k] not in valid_stock:
            valid_stock.append(data_category["標題三十七"][k])

    new_stock_list = []
    for stock in valid_stock:
        if stock not in stock_list:
            new_stock_list.append(stock)
            
    new_stock = []
    for stock in new_stock_list:
        try:
            xls_df_day = pd.read_csv("C:\Users\User\Desktop\Daily Report\"+stock+\"-day.csv",engine='python', encoding = "big5",index_col=0)
        except:
            continue
        
        for k in range(len(xls_df_day)):
            if int(xls_df_day["日期"][k]) == int(date):        
                new_stock.append([int(date),stock,xls_df_day["日開盤價"][k],xls_df_day["日收盤價"][k],0,xls_df_day["漲跌幅"][k],xls_df_day["漲跌幅"][k],0,xls_df_day["日最高價"][k],1,"",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""])
                break
            
    new_stock = pd.DataFrame(new_stock,columns = ["候選日期","標題三十七","開盤價","收盤價","當日收盤價","漲跌幅","累計漲跌幅","跌破5日均線次數","突破最高價","候選天數","乖離率","第一天k棒","第一天","第二天k棒","第二天","第三天k棒","第三天","第四天k棒","第四天","第五天k棒","第五天","第六天k棒","第六天","第七天k棒","第七天","第八天k棒","第八天","第九天k棒","第九天","第十天k棒","第十天","第十一天k棒","第十一天","第十二天k棒","第十二天","第十三天k棒","第十三天","第十四天k棒","第十四天","第十五天k棒","第十五天","標題十-1","標題六十-1","標題六十三-1","標題六十一-1","標題六十二-1","標題六十四-1","標題六十五-1","標題十-2","標題六十-2","標題六十三-2","標題六十一-2","標題六十二-2","標題六十四-2","標題六十五-2","標題十-3","標題六十-3","標題六十三-3","標題六十一-3","標題六十二-3","標題六十四-3","標題六十五-3","標題十-4","標題六十-4","標題六十三-4","標題六十一-4","標題六十二-4","標題六十四-4","標題六十五-4","標題十-5","標題六十-5","標題六十三-5","標題六十一-5","標題六十二-5","標題六十四-5","標題六十五-5","標題十-6","標題六十-6","標題六十三-6","標題六十一-6","標題六十二-6","標題六十四-6","標題六十五-6","標題十-7","標題六十-7","標題六十三-7","標題六十一-7","標題六十二-7","標題六十四-7","標題六十五-7","標題十-8","標題六十-8","標題六十三-8","標題六十一-8","標題六十二-8","標題六十四-8","標題六十五-8","標題十-9","標題六十-9","標題六十三-9","標題六十一-9","標題六十二-9","標題六十四-9","標題六十五-9","標題十-10","標題六十-10","標題六十三-10","標題六十一-10","標題六十二-10","標題六十四-10","標題六十五-10","股數","股數_1"] )    

    for loca in range(len(new_stock["標題三十七"])):
        time = 0
        for k in range(len(data_category)):
            if data_category["標題三十七"][k] == new_stock["標題三十七"][loca]:
                time = time + 1
                if time == 1:
                    new_stock["標題十-1"][loca] = data_category["標題十"][k]
                    new_stock["標題六十-1"][loca] = data_category["標題六十"][k]
                    new_stock["標題六十三-1"][loca] = data_category["標題六十三"][k]
                    new_stock["標題六十一-1"][loca] = data_category["標題六十一"][k]
                    new_stock["標題六十二-1"][loca] = data_category["標題六十二"][k]
                    new_stock["標題六十五-1"][loca] = data_category["標題六十五"][k]
                    new_stock["標題六十四-1"][loca] = data_category["標題六十四"][k]
                elif time == 2:                
                    new_stock["標題十-2"][loca] = data_category["標題十"][k]
                    new_stock["標題六十-2"][loca] = data_category["標題六十"][k]
                    new_stock["標題六十三-2"][loca] = data_category["標題六十三"][k]
                    new_stock["標題六十一-2"][loca] = data_category["標題六十一"][k]
                    new_stock["標題六十二-2"][loca] = data_category["標題六十二"][k]                 
                    new_stock["標題六十五-2"][loca] = data_category["標題六十五"][k]
                    new_stock["標題六十四-2"][loca] = data_category["標題六十四"][k]
                elif time == 3:                 
                    new_stock["標題十-3"][loca] = data_category["標題十"][k]
                    new_stock["標題六十-3"][loca] = data_category["標題六十"][k]
                    new_stock["標題六十三-3"][loca] = data_category["標題六十三"][k]
                    new_stock["標題六十一-3"][loca] = data_category["標題六十一"][k]
                    new_stock["標題六十二-3"][loca] = data_category["標題六十二"][k]
                    new_stock["標題六十五-3"][loca] = data_category["標題六十五"][k]
                    new_stock["標題六十四-3"][loca] = data_category["標題六十四"][k]
                elif time == 4:        
                    new_stock["標題十-4"][loca] = data_category["標題十"][k]
                    new_stock["標題六十-4"][loca] = data_category["標題六十"][k]
                    new_stock["標題六十三-4"][loca] = data_category["標題六十三"][k]
                    new_stock["標題六十一-4"][loca] = data_category["標題六十一"][k]
                    new_stock["標題六十二-4"][loca] = data_category["標題六十二"][k]
                    new_stock["標題六十五-4"][loca] = data_category["標題六十五"][k]
                    new_stock["標題六十四-4"][loca] = data_category["標題六十四"][k]
                elif time == 5:  
                    new_stock["標題十-5"][loca] = data_category["標題十"][k]
                    new_stock["標題六十-5"][loca] = data_category["標題六十"][k]
                    new_stock["標題六十三-5"][loca] = data_category["標題六十三"][k]
                    new_stock["標題六十一-5"][loca] = data_category["標題六十一"][k]
                    new_stock["標題六十二-5"][loca] = data_category["標題六十二"][k]
                    new_stock["標題六十五-5"][loca] = data_category["標題六十五"][k]
                    new_stock["標題六十四-5"][loca] = data_category["標題六十四"][k]
                new_stock["股數"][loca] = len(list(set(data_category["標題三十七"].tolist())))
                new_stock["股數_1"][loca] = data_category["標題六十"][len(data_category)-1]
                
    xls_df_strong = xls_df_strong.append(new_stock,ignore_index=True)

    for k in range(len(xls_df_strong)):
        if int(xls_df_strong["候選日期"][k]) == int(date):
            continue
        stock = xls_df_strong["標題三十七"][k]
        for i in range(len(data_category)):
            if data_category["標題三十七"][i] == stock :
                category_name = data_category["標題十"][i]
                for column in xls_df_strong.columns:
                    if xls_df_strong[column][k] == category_name:
                        switch = 0
                        if len(list(set(data_category["標題三十七"].tolist()))) < 90 and data_category["標題六十"][len(data_category)-1] < 80 :
                            if 0 < data_category["標題六十"][i] <= 10 and 0 < data_category["標題六十二"][i] <= 10 :  
                                switch = 1
                            elif 0 < data_category["標題六十"][i] <= 10 and 0 < data_category["標題六十一"][i] <= 10 :  
                                switch = 1
                            elif 0 < data_category["標題六十三"][i] <= 10 and 0 < data_category["標題六十二"][i] <= 10 :
                                switch = 1 
                            elif 0 < data_category["標題六十三"][i] <= 10 and 0 < data_category["標題六十一"][i] <= 10 :  
                                switch = 1        
                            elif 0 < data_category["標題六十"][i] <= 12 and 0 < data_category["標題六十二"][i] <= 3 :  
                                switch = 1     
                            elif 0 < data_category["標題六十"][i] <= 12 and 0 < data_category["標題六十一"][i] <= 3 :  
                                switch = 1   
                            elif 0 < data_category["標題六十三"][i] <= 12 and 0 < data_category["標題六十二"][i] <= 3 :
                                switch = 1 
                            elif 0 < data_category["標題六十三"][i] <= 12 and 0 < data_category["標題六十一"][i] <= 3 :  
                                switch = 1
                                    
                        elif len(list(set(data_category["標題三十七"].tolist()))) < 90 and data_category["標題六十"][len(data_category)-1] >= 80 :
                            if 0 < data_category["標題六十"][i] <= 10 and 0 < data_category["標題六十二"][i] <= 20 :  
                                switch = 1
                            elif 0 < data_category["標題六十"][i] <= 10 and 0 < data_category["標題六十一"][i] <= 20 :  
                                switch = 1
                            elif 0 < data_category["標題六十三"][i] <= 10 and 0 < data_category["標題六十二"][i] <= 20 :
                                switch = 1 
                            elif 0 < data_category["標題六十三"][i] <= 10 and 0 < data_category["標題六十一"][i] <= 20 :  
                                switch = 1        
                            elif 0 < data_category["標題六十"][i] <= 12 and 0 < data_category["標題六十二"][i] <= 6 :  
                                switch = 1     
                            elif 0 < data_category["標題六十"][i] <= 12 and 0 < data_category["標題六十一"][i] <= 6 :  
                                switch = 1   
                            elif 0 < data_category["標題六十三"][i] <= 12 and 0 < data_category["標題六十二"][i] <= 6 :
                                switch = 1 
                            elif 0 < data_category["標題六十三"][i] <= 12 and 0 < data_category["標題六十一"][i] <= 6 :  
                                switch = 1    
                                
                        elif len(list(set(data_category["標題三十七"].tolist()))) >= 90 and data_category["標題六十"][len(data_category)-1] < 80 :
                            if 0 < data_category["標題六十"][i] <= 20 and 0 < data_category["標題六十二"][i] <= 10 :  
                                switch = 1
                            elif 0 < data_category["標題六十"][i] <= 20 and 0 < data_category["標題六十一"][i] <= 10 :  
                                switch = 1
                            elif 0 < data_category["標題六十三"][i] <= 20 and 0 < data_category["標題六十二"][i] <= 10 :
                                switch = 1 
                            elif 0 < data_category["標題六十三"][i] <= 20 and 0 < data_category["標題六十一"][i] <= 10 :  
                                switch = 1                 
                            
                        elif len(list(set(data_category["標題三十七"].tolist()))) >= 90 and data_category["標題六十"][len(data_category)-1] >= 80 :
                            if 0 < data_category["標題六十"][i] <= 20 and 0 < data_category["標題六十二"][i] <= 20 :  
                                switch = 1
                            elif 0 < data_category["標題六十"][i] <= 20 and 0 < data_category["標題六十一"][i] <= 20 :  
                                switch = 1
                            elif 0 < data_category["標題六十三"][i] <= 20 and 0 < data_category["標題六十二"][i] <= 20 :
                                switch = 1 
                            elif 0 < data_category["標題六十三"][i] <= 20 and 0 < data_category["標題六十一"][i] <= 20 :  
                                switch = 1     
                        
                        if switch == 1 :
                            xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+1]][k] = data_category["標題六十"][i]
                            xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+2]][k] = data_category["標題六十三"][i]    
                            xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+3]][k] = data_category["標題六十一"][i]
                            xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+4]][k] = data_category["標題六十二"][i]
                            xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+5]][k] = data_category["標題六十四"][i]
                            xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+6]][k] = data_category["標題六十五"][i]
        
                        break
                    elif column == xls_df_strong.columns[-1]:
                        for column in xls_df_strong.columns:
                            if xls_df_strong[column][k] == "" and column != "乖離率" and column != "跌破5日隔天站回":
                                xls_df_strong[column][k] = data_category["標題十"][i]
                                xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+1]][k] = data_category["標題六十"][i]
                                xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+2]][k] = data_category["標題六十三"][i]
                                xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+3]][k] = data_category["標題六十一"][i]
                                xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+4]][k] = data_category["標題六十二"][i]
                                xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+5]][k] = data_category["標題六十四"][i]
                                xls_df_strong[xls_df_strong.columns[np.array(xls_df_strong.columns).tolist().index(column)+6]][k] = data_category["標題六十五"][i]   
                                break
                        
    for k in range(len(xls_df_strong)):
        if int(xls_df_strong["候選日期"][k]) == int(date):
            continue
        if int(xls_df_strong["候選日期"][k]) <= delete_date:
            xls_df_strong.drop([k],axis = 0, inplace = True)    
            continue
        stock = xls_df_strong["標題三十七"][k]
        
        xls_df_day = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)
        for i in range(len(xls_df_day)):
            if int(xls_df_day["日期"][i]) == int(date):
                break

        xls_df_strong["候選天數"][k] = xls_df_strong["候選天數"][k] + 1
        xls_df_strong["當日收盤價"][k] = xls_df_day["日收盤價"][i]
        xls_df_strong["突破最高價"][k] = round((xls_df_day["日收盤價"][i] - max(xls_df_day["日最高價"][i-5:i])) / max(xls_df_day["日最高價"][i-5:i]) * 100,2)
        xls_df_strong["乖離率"][k] = round((xls_df_day["日收盤價"][i] - xls_df_day["5日均線"][i]) / xls_df_day["5日均線"][i] * 100,2)
        xls_df_strong["累計漲跌幅"][k] = round((xls_df_day["日收盤價"][i] - xls_df_strong["開盤價"][k]) / xls_df_strong["開盤價"][k] * 100,2)
        for column in xls_df_strong.columns:
            if "第" in column and "k" in column and xls_df_strong[column][k] == 0 and xls_df_strong["當日收盤價"][k] >= xls_df_strong["收盤價"][k] :
                xls_df_strong[column][k] = "上"
                continue
            elif "第" in column and "k" in column and xls_df_strong[column][k] == 0 and xls_df_strong["收盤價"][k] > xls_df_strong["當日收盤價"][k] >= xls_df_strong["開盤價"][k] :
                xls_df_strong[column][k] = "內"
                continue
            if "第" in column and "k" not in column and xls_df_strong[column][k] == 0:
                xls_df_strong[column][k] = xls_df_day["漲跌幅"][i]
                break
        
        sum = 0
        for p in range(i-3,i+1):
            sum = xls_df_day["日收盤價"][p] + sum

        if xls_df_day["日最低價"][i] < xls_df_day["5日均線"][i]*0.997:
            xls_df_strong["跌破5日均線次數"][k] = xls_df_strong["跌破5日均線次數"][k] + 1
            if xls_df_strong["跌破5日均線次數"][k] == 2 :
                xls_df_strong.drop([k],axis = 0, inplace = True)               

    xls_df_strong.index = range(len(xls_df_strong))                
    if switch_big == 0 :  
        xls_df_strong.to_excel(r"C:\Users\User\Desktop\每日計算報告\強勢股列表-上市"+str(date)+".xls")    
        print(str(date)+"強勢股列表-上市完成")   
    elif switch_big == 1 :  
        xls_df_strong.to_excel(r"C:\Users\User\Desktop\每日計算報告\強勢股列表-上櫃"+str(date)+".xls")    
        print(str(date)+"強勢股列表-上櫃完成")   

    trade_list,candidate_list = [],[]    
    for k in range(len(data_category)):         
        if data_category["標題六十"][len(data_category)-1] < 90 and len(list(set(data_category["標題三十七"].tolist()))) < 80 :
            if data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十"][k] <= 10 and data_category["標題六十一"][k] <= 10 :     
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十"][k] <= 10 and data_category["標題六十一"][k] <= 20 and data_category["標題六十二"][k] <= 10:   
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十三"][k] <= 10 and data_category["標題六十一"][k] <= 10 :     
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十三"][k] <= 10 and data_category["標題六十一"][k] <= 20 and data_category["標題六十二"][k] <= 10:   
                candidate_list.append(data_category["標題三十七"][k])   
                
        elif data_category["標題六十"][len(data_category)-1] < 90 and len(list(set(data_category["標題三十七"].tolist()))) >= 80 :
            if data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十"][k] <= 10 and data_category["標題六十一"][k] <= 20 :     
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十"][k] <= 10 and data_category["標題六十一"][k] <= 30 and data_category["標題六十二"][k] <= 20:   
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十三"][k] <= 10 and data_category["標題六十一"][k] <= 20 :     
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十三"][k] <= 10 and data_category["標題六十一"][k] <= 30 and data_category["標題六十二"][k] <= 20:   
                candidate_list.append(data_category["標題三十七"][k])   
                
        elif data_category["標題六十"][len(data_category)-1] >= 90 and len(list(set(data_category["標題三十七"].tolist()))) < 80 :
            if data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十"][k] <= 20 and data_category["標題六十一"][k] <= 10 :     
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十"][k] <= 20 and data_category["標題六十一"][k] <= 20 and data_category["標題六十二"][k] <= 10:   
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十三"][k] <= 20 and data_category["標題六十一"][k] <= 10 :     
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十三"][k] <= 20 and data_category["標題六十一"][k] <= 20 and data_category["標題六十二"][k] <= 10:   
                candidate_list.append(data_category["標題三十七"][k])                 
                
        elif data_category["標題六十"][len(data_category)-1] >= 90 and len(list(set(data_category["標題三十七"].tolist()))) >= 80 :
            if data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十"][k] <= 20 and data_category["標題六十一"][k] <= 20 :     
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十"][k] <= 20 and data_category["標題六十一"][k] <= 30 and data_category["標題六十二"][k] <= 20:   
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十三"][k] <= 20 and data_category["標題六十一"][k] <= 20 :     
                candidate_list.append(data_category["標題三十七"][k])
            elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十三"][k] <= 20 and data_category["標題六十一"][k] <= 30 and data_category["標題六十二"][k] <= 20:   
                candidate_list.append(data_category["標題三十七"][k])                   
                
    candidate_list = list(set(candidate_list))
        
    for stock in candidate_list:
        stratgy1,stratgy2,stratgy3,stratgy4,stratgy5,stratgy6,stratgy7,stratgy8,stratgy9,stratgy10,stratgy11,stratgy12,stratgy13,stratgy14,stratgy15,stratgy16 = 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
        rank_list1,rank_list2,rank_list3,rank_list4,rank_list5,rank_list6,rank_list7,rank_list8,rank_list9,rank_list10,rank_list11,rank_list12,rank_list13,rank_list14,rank_list15,rank_list16, = ["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""],["","","","","","","","","",""]
   
        for k in range(len(data_category)):
            if data_category["標題三十七"][k] == stock: 
                if data_category["標題六十"][len(data_category)-1] < 90 and len(list(set(data_category["標題三十七"].tolist()))) < 80 :
                    if data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十"][k] <= 10 and data_category["標題六十一"][k] <= 10 :
                        stratgy1 = 1
                        if rank_list1[5] != "" and rank_list1[6] != rank_list1[5]:    
                            continue
                        else:
                            rank_list1 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                    elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十"][k] <= 10 and data_category["標題六十一"][k] <= 20 and data_category["標題六十二"][k] <= 10:   
                        stratgy2 = 1
                        if rank_list2[5] != "" and rank_list2[6] != rank_list2[5]:    
                            continue
                        else:
                            rank_list2 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                    elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十三"][k] <= 10 and data_category["標題六十一"][k] <= 10 :     
                        stratgy3 = 1
                        if rank_list3[5] != "" and rank_list3[6] != rank_list3[5]:    
                            continue
                        else:
                            rank_list3 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                    elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十三"][k] <= 10 and data_category["標題六十一"][k] <= 20 and data_category["標題六十二"][k] <= 10:   
                        stratgy4 = 1
                        if rank_list4[5] != "" and rank_list4[6] != rank_list4[5]:    
                            continue
                        else:
                            rank_list4 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                        
                elif data_category["標題六十"][len(data_category)-1] < 90 and len(list(set(data_category["標題三十七"].tolist()))) >= 80 :
                    if data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十"][k] <= 10 and data_category["標題六十一"][k] <= 20 :     
                        stratgy5 = 1
                        if rank_list5[5] != "" and rank_list5[6] != rank_list5[5]:    
                            continue
                        else:       
                            rank_list5 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                    elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十"][k] <= 10 and data_category["標題六十一"][k] <= 30 and data_category["標題六十二"][k] <= 20:   
                        stratgy6 = 1
                        if rank_list6[5] != "" and rank_list6[6] != rank_list6[5]:    
                            continue
                        else:
                            rank_list6 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                    elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十三"][k] <= 10 and data_category["標題六十一"][k] <= 20 :     
                        stratgy7 = 1
                        if rank_list7[5] != "" and rank_list7[6] != rank_list7[5]:    
                            continue
                        else:
                            rank_list7 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                    elif data_category["標題六十七"][k] <= 10 and 0 < data_category["標題六十三"][k] <= 10 and data_category["標題六十一"][k] <= 30 and data_category["標題六十二"][k] <= 20:   
                        stratgy8 = 1
                        if rank_list8[5] != "" and rank_list8[6] != rank_list8[5]:    
                            continue
                        else:
                            rank_list8 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                elif data_category["標題六十"][len(data_category)-1] >= 90 and len(list(set(data_category["標題三十七"].tolist()))) < 80 :
                    if data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十"][k] <= 20 and data_category["標題六十一"][k] <= 10 :     
                        stratgy9 = 1
                        if rank_list9[5] != "" and rank_list9[6] != rank_list9[5]:    
                            continue
                        else:
                            rank_list9 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                    elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十"][k] <= 20 and data_category["標題六十一"][k] <= 20 and data_category["標題六十二"][k] <= 10:   
                        stratgy10 = 1
                        if rank_list10[5] != "" and rank_list10[6] != rank_list10[5]:    
                            continue
                        else:
                            rank_list10 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""]

                    elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十三"][k] <= 20 and data_category["標題六十一"][k] <= 10 :     
                        stratgy11 = 1
                        if rank_list11[5] != "" and rank_list11[6] != rank_list11[5]:    
                            continue
                        else:
                            rank_list11 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                    elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十三"][k] <= 20 and data_category["標題六十一"][k] <= 20 and data_category["標題六十二"][k] <= 10:   
                        stratgy12 = 1
                        if rank_list12[5] != "" and rank_list12[6] != rank_list12[5]:    
                            continue
                        else:
                            rank_list12 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""]              

                        
                elif data_category["標題六十"][len(data_category)-1] >= 90 and len(list(set(data_category["標題三十七"].tolist()))) >= 80 :
                    if data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十"][k] <= 20 and data_category["標題六十一"][k] <= 20 :     
                        stratgy13 = 1
                        if rank_list13[5] != "" and rank_list13[6] != rank_list13[5]:    
                            continue
                        else:
                            rank_list13 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""]  

                    elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十"][k] <= 20 and data_category["標題六十一"][k] <= 30 and data_category["標題六十二"][k] <= 20:   
                        stratgy14 = 1
                        if rank_list14[5] != "" and rank_list14[6] != rank_list14[5]:    
                            continue
                        else:
                            rank_list14 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 

                    elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十三"][k] <= 20 and data_category["標題六十一"][k] <= 20 :     
                        stratgy15 = 1
                        if rank_list15[5] != "" and rank_list15[6] != rank_list15[5]:    
                            continue
                        else:
                            rank_list15 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""]  

                    elif data_category["標題六十七"][k] <= 20 and 0 < data_category["標題六十三"][k] <= 20 and data_category["標題六十一"][k] <= 30 and data_category["標題六十二"][k] <= 20:   
                        stratgy16 = 1
                        if rank_list16[5] != "" and rank_list16[6] != rank_list16[5]:    
                            continue
                        else:
                            rank_list16 = [int(data_category["標題六十"][k]),int(data_category["標題六十三"][k]),int(data_category["標題六十一"][k]),int(data_category["標題六十二"][k]),int(data_category["標題六十七"][k]),data_category["標題九"][k],data_category["標題十"][k],data_category["標題六十五"][k],data_category["標題六十四"][k],""] 


        smallest = []
        for list_ in [rank_list1,rank_list2,rank_list3,rank_list4,rank_list5,rank_list6,rank_list7,rank_list8,rank_list9,rank_list10,rank_list11,rank_list12,rank_list13,rank_list14,rank_list15,rank_list16]:
            if list_ == ["","","","","","","","","",""]:
                continue
            if smallest == [] :
                smallest = [list_[8],list_[7]]
            elif smallest[0] > list_[8]:
                smallest = [list_[8],list_[7]]
            elif smallest[0] == list_[8] and smallest[1] > list_[7]:
                smallest = [list_[8],list_[7]]

        total_stock = len(list(set(data_category["標題三十七"].tolist())))
        total_category = data_category["標題六十"][len(data_category)-1]
                                                
        trade_list.append([date,stock,rank_list1[0],rank_list1[1],rank_list1[2],rank_list1[3],rank_list1[4],rank_list1[5],rank_list1[6],rank_list1[7],rank_list1[8],rank_list1[9],rank_list2[0],rank_list2[1],rank_list2[2],rank_list2[3],rank_list2[4],rank_list2[5],rank_list2[6],rank_list2[7],rank_list2[8],rank_list2[9],rank_list3[0],rank_list3[1],rank_list3[2],rank_list3[3],rank_list3[4],rank_list3[5],rank_list3[6],rank_list3[7],rank_list3[8],rank_list3[9],rank_list4[0],rank_list4[1],rank_list4[2],rank_list4[3],rank_list4[4],rank_list4[5],rank_list4[6],rank_list4[7],rank_list4[8],rank_list4[9],rank_list5[0],rank_list5[1],rank_list5[2],rank_list5[3],rank_list5[4],rank_list5[5],rank_list5[6],rank_list5[7],rank_list5[8],rank_list5[9],rank_list6[0],rank_list6[1],rank_list6[2],rank_list6[3],rank_list6[4],rank_list6[5],rank_list6[6],rank_list6[7],rank_list6[8],rank_list6[9],rank_list7[0],rank_list7[1],rank_list7[2],rank_list7[3],rank_list7[4],rank_list7[5],rank_list7[6],rank_list7[7],rank_list7[8],rank_list7[9],rank_list8[0],rank_list8[1],rank_list8[2],rank_list8[3],rank_list8[4],rank_list8[5],rank_list8[6],rank_list8[7],rank_list8[8],rank_list8[9],rank_list9[0],rank_list9[1],rank_list9[2],rank_list9[3],rank_list9[4],rank_list9[5],rank_list9[6],rank_list9[7],rank_list9[8],rank_list9[9],rank_list10[0],rank_list10[1],rank_list10[2],rank_list10[3],rank_list10[4],rank_list10[5],rank_list10[6],rank_list10[7],rank_list10[8],rank_list10[9],rank_list11[0],rank_list11[1],rank_list11[2],rank_list11[3],rank_list11[4],rank_list11[5],rank_list11[6],rank_list11[7],rank_list11[8],rank_list11[9],rank_list12[0],rank_list12[1],rank_list12[2],rank_list12[3],rank_list12[4],rank_list12[5],rank_list12[6],rank_list12[7],rank_list12[8],rank_list12[9],rank_list13[0],rank_list13[1],rank_list13[2],rank_list13[3],rank_list13[4],rank_list13[5],rank_list13[6],rank_list13[7],rank_list13[8],rank_list13[9],rank_list14[0],rank_list14[1],rank_list14[2],rank_list14[3],rank_list14[4],rank_list14[5],rank_list14[6],rank_list14[7],rank_list14[8],rank_list14[9],rank_list15[0],rank_list15[1],rank_list15[2],rank_list15[3],rank_list15[4],rank_list15[5],rank_list15[6],rank_list15[7],rank_list15[8],rank_list15[9],rank_list16[0],rank_list16[1],rank_list16[2],rank_list16[3],rank_list16[4],rank_list16[5],rank_list16[6],rank_list16[7],rank_list16[8],rank_list16[9],stratgy1,stratgy2,stratgy3,stratgy4,stratgy5,stratgy6,stratgy7,stratgy8,stratgy9,stratgy10,stratgy11,stratgy12,stratgy13,stratgy14,stratgy15,stratgy16,smallest[0],smallest[1],total_stock,total_category])

    candidate_list = []   
    for k in range(len(xls_df_strong)):            
        column_list = []
        for column in xls_df_strong.columns[11:]:
            if xls_df_strong[column][k] != 0 :
                column_list.append(column)
            else:
                break

        for column in column_list:
            if "k棒" not in column and abs(xls_df_strong[column][k]) < xls_df_strong["漲跌幅"][k]/2 and abs(xls_df_strong["乖離率"][k]) <= 1 :
                candidate_list.append(xls_df_strong["標題三十七"][k])  
                break
            elif "k棒" not in column and xls_df_strong["突破最高價"][k] >= 0 :
                candidate_list.append(xls_df_strong["標題三十七"][k])                  
                break 
            
    candidate_list = list(set(candidate_list))    
    candidate_list_qualified = []
    for stock in candidate_list:
        switch = 0
        for k in range(len(xls_df_strong)):   
            if xls_df_strong["標題三十七"][k] == stock:
                if xls_df_strong["股數"][k] < 90 and xls_df_strong["股數_1"][k] < 80 :
                    for num in range(1,11):
                        if 0 < xls_df_strong["標題六十-"+str(num)][k] <= 10 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 10 :  
                            switch = 1
                        elif 0 < xls_df_strong["標題六十-"+str(num)][k] <= 10 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 10 :  
                            switch = 1
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 10 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 10 :
                            switch = 1 
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 10 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 10 :  
                            switch = 1        
                        elif 0 < xls_df_strong["標題六十-"+str(num)][k] <= 12 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 3 :  
                            switch = 1     
                        elif 0 < xls_df_strong["標題六十-"+str(num)][k] <= 12 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 3 :  
                            switch = 1   
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 12 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 3 :
                            switch = 1 
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 12 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 3 :  
                            switch = 1
                            
                elif xls_df_strong["股數"][k] < 90 and xls_df_strong["股數_1"][k] >= 80 :
                    for num in range(1,11):
                        if 0 < xls_df_strong["標題六十-"+str(num)][k] <= 10 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 20 :  
                            switch = 1
                        elif 0 < xls_df_strong["標題六十-"+str(num)][k] <= 10 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 20 :  
                            switch = 1
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 10 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 20 :
                            switch = 1 
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 10 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 20 :  
                            switch = 1        
                        elif 0 < xls_df_strong["標題六十-"+str(num)][k] <= 12 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 6 :  
                            switch = 1     
                        elif 0 < xls_df_strong["標題六十-"+str(num)][k] <= 12 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 6 :  
                            switch = 1   
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 12 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 6 :
                            switch = 1 
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 12 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 6 :  
                            switch = 1    
                        
                elif xls_df_strong["股數"][k] >= 90 and xls_df_strong["股數_1"][k] < 80 :
                    for num in range(1,11):
                        if 0 < xls_df_strong["標題六十-"+str(num)][k] <= 20 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 10 :  
                            switch = 1
                        elif 0 < xls_df_strong["標題六十-"+str(num)][k] <= 20 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 10 :  
                            switch = 1
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 20 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 10 :
                            switch = 1 
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 20 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 10 :  
                            switch = 1                 
                        
                elif xls_df_strong["股數"][k] >= 90 and xls_df_strong["股數_1"][k] >= 80 :
                    for num in range(1,11):
                        if 0 < xls_df_strong["標題六十-"+str(num)][k] <= 20 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 20 :  
                            switch = 1
                        elif 0 < xls_df_strong["標題六十-"+str(num)][k] <= 20 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 20 :  
                            switch = 1
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 20 and 0 < xls_df_strong["標題六十二-"+str(num)][k] <= 20 :
                            switch = 1 
                        elif 0 < xls_df_strong["標題六十三-"+str(num)][k] <= 20 and 0 < xls_df_strong["標題六十一-"+str(num)][k] <= 20 :  
                            switch = 1        
     
                if switch == 1 :
                    candidate_list_qualified.append(stock)
                    break
                
    for stock in candidate_list_qualified:

        smallest = []        
        for k in range(len(xls_df_strong)):
            if xls_df_strong["標題三十七"][k] == stock:
                for num in range(1,11):
                    if xls_df_strong["標題六十四-"+str(num)][k] == "" :
                        continue
                    if num == 1 :
                        smallest = [xls_df_strong["標題六十四-"+str(num)][k], xls_df_strong["標題六十五-"+str(num)][k]]
                    elif smallest[0] > xls_df_strong["標題六十四-"+str(num)][k] :
                        smallest = [xls_df_strong["標題六十四-"+str(num)][k], xls_df_strong["標題六十五-"+str(num)][k]]
                    elif smallest[0] == xls_df_strong["標題六十四-"+str(num)][k] and smallest[1] > xls_df_strong["標題六十五-"+str(num)][k]:
                        smallest = [xls_df_strong["標題六十四-"+str(num)][k], xls_df_strong["標題六十五-"+str(num)][k]]
                break
                     
        trade_list.append([date,stock,"","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",smallest[0],smallest[1],"",""])
                
    trade_list = pd.DataFrame(trade_list,columns = ["日期","標題三十七","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","標題六十","標題六十三","標題六十一","標題六十二","標題六十七","標題九","標題十","標題六十五","標題六十四","相同名稱","策略1","策略2","策略3","策略4","策略5","策略6","策略7","策略8","策略9","策略10","策略11","策略12","策略13","策略14","策略15","策略16","排名","排名_1","股數","股數_1"])
    if switch_big == 0 :  
        trade_list.to_excel("C:\Users\User\Desktop\每日計算報告\市候選"+str(date)+".xlsx")
    elif switch_big == 1 :   
        trade_list.to_excel("C:\Users\User\Desktop\每日計算報告\櫃候選"+str(date)+".xlsx")                         