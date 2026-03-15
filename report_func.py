# -*- codinD: utf-8 -*-
"""
Created on Thu Dec 13 20:45:12 2018

@author: YAO
"""

import datetime
import pandas as pd
import daily_count_func
import numpy as np
import price_slope_daily_new
import math

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

def excel_locate(xls_df_basic,stock):
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
    xls_file_big_d.index = range(xls_df_report.index[-1]-len(xls_file_big_d)+1,xls_df_report.index[-1]+1)
    
    return xls_file_big_d

def erase_for_big_w(xls_df_big_w,xls_df_week):
    xls_df_date = set(list(xls_df_week["日期"]))
    xls_df_big_date = set(list(xls_df_big_w["日期"]))
    diff = xls_df_big_date.difference(xls_df_date)
    diff = list(diff)   
    diff.sort(reverse = False)

    for i in range(len(diff)):
        for j in range(len(xls_df_week["日期"])):          
            if be_format(str(int(xls_df_week["日期"][j])))[0] == be_format(str(int(diff[i])))[0]:
                if be_format(str(int(xls_df_week["日期"][j])))[1] == be_format(str(int(diff[i])))[1]:
                    xls_df_week["日期"][j] = diff[i]
                    break

    xls_df_date = set(list(xls_df_week["日期"]))
    xls_df_big_date = set(list(xls_df_big_w["日期"]))
    diff = xls_df_big_date.difference(xls_df_date)
    diff = list(diff)   
    diff.sort(reverse = False)

    for i in range(xls_df_big_w.index[0],xls_df_big_w.index[-1]+1):
        if xls_df_big_w["日期"][i] in diff :
            xls_df_big_w.drop([i],axis = 0, inplace = True)
    xls_df_big_w.index = range(xls_df_week.index[-1]-len(xls_df_big_w)+1,xls_df_week.index[-1]+1)
    
    return xls_df_big_w

################################
# 主要運行程式碼-每日報表_新-日 ####   
###############################
def daily_count(newdata,stock,xls_df_multiply,switch_big):
                                    
    if type(newdata["收盤價"][0]) == str:
        if "-" in newdata["收盤價"][0].strip():
            print(stock+"日報表-每日行情 不需變動")
            print(str(datetime.datetime.now()))      
            return None                            

    xls_df_report = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)

    date_report = set(list(xls_df_report["日期"])) 
    date_df = set(list(newdata["日期"]))              
    date_list = list(date_df.difference(date_report))
    date_list.sort(reverse = False)  
    switch = 0
    for data_date in date_list: 
        if data_date < int(np.array(xls_df_report["日期"]).tolist()[-1]) :
            continue
        elif type(newdata["收盤價"][np.array(newdata["日期"]).tolist().index(data_date)]) == str:
            continue
        data_list = []
        for loca in range(len(date_df)):
            if int(newdata["日期"][loca]) == int(data_date):
                try:
                    data_list.append([newdata["日期"][loca],newdata["成交股數"][loca],newdata["成交筆數"][loca],newdata["成交金額"][loca],newdata["開盤價"][loca],newdata["最高價"][loca],newdata["最低價"][loca],newdata["收盤價"][loca]])
                    for num in range(len(xls_df_report.columns)-len(data_list[0])):
                        data_list[0].append(0.0001)
                    df = pd.DataFrame(data_list,columns= xls_df_report.columns )
                    xls_df_report = pd.concat([xls_df_report,df],axis = 0) 
                    xls_df_report.index = range(len(xls_df_report)) 
                    xls_df_report["5日均線"][len(xls_df_report)-1] = daily_count_func.ma_day(xls_df_report,5)       
                    xls_df_report["10日均線"][len(xls_df_report)-1] = daily_count_func.ma_day(xls_df_report,10)
                    xls_df_report["均線"][len(xls_df_report)-1] = daily_count_func.ma_day(xls_df_report,20) 
                    xls_df_report["60日均線"][len(xls_df_report)-1] = daily_count_func.ma_day(xls_df_report,60) 
                    xls_df_report["5日均量"][len(xls_df_report)-1] = daily_count_func.volume_ma_day(xls_df_report,5)
                    xls_df_report["日EMA20"][len(xls_df_report)-1] = daily_count_func.day_moving_average(xls_df_report)[0]
                    xls_df_report["參考指標"][len(xls_df_report)-1] = daily_count_func.day_moving_average(xls_df_report)[2]
                    xls_df_report["參考指標"][len(xls_df_report)-1] = daily_count_func.day_moving_average(xls_df_report)[3]
                
                    xls_df_report["日上軌"][len(xls_df_report)-1] = daily_count_func.day_bbands(xls_df_report)[0]
                    xls_df_report["日中軸"][len(xls_df_report)-1] = daily_count_func.day_bbands(xls_df_report)[1]     
                    xls_df_report["日下軌"][len(xls_df_report)-1] = daily_count_func.day_bbands(xls_df_report)[2]             
                    xls_df_report["參考指標_87"][len(xls_df_report)-1] = daily_count_func.day_ma_angle(xls_df_report)[0]    
                    xls_df_report["參考指標_88"][len(xls_df_report)-1] = daily_count_func.day_ma_angle(xls_df_report)[4]    
                    xls_df_report["參考指標_89"][len(xls_df_report)-1] = daily_count_func.angle20_ma_day(xls_df_report,5)              
                    xls_df_report["參考指標_90"][len(xls_df_report)-1] = daily_count_func.angle20_ma_day(xls_df_report,10) 
                    xls_df_report["參考指標_91"][len(xls_df_report)-1] = daily_count_func.angle20_ema_day(xls_df_report,5)              
                    xls_df_report["參考指標_92"][len(xls_df_report)-1] = daily_count_func.angle20_ema_day(xls_df_report,10) 
                    
                    xls_df_report["參考指標_93"][len(xls_df_report)-1] = daily_count_func.day(xls_df_report)[0] 
                    xls_df_report["參考指標_94"][len(xls_df_report)-1] = daily_count_func.day(xls_df_report)[1] 
                    xls_df_report["參考指標_95"][len(xls_df_report)-1] = daily_count_func.day(xls_df_report)[2] 
                    xls_df_report["參考指標_96"][len(xls_df_report)-1] = daily_count_func.day(xls_df_report)[3] 
                    xls_df_report["參考指標_97"][len(xls_df_report)-1] = daily_count_func.day(xls_df_report)[4]  
                    
                    xls_df_report["漲跌幅"][len(xls_df_report)-1] = daily_count_func.day_down_up_percentage(xls_df_report)
    
                    switch = 1
                    break
                except:
                    print(stock+"日報表-每日行情 計算失敗")
                    print(str(datetime.datetime.now()))
                    if switch_big == 0:   
                        xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
                    if switch_big == 1:   
                        xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-櫃.xlsx")                    
                    xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0])         
                    for i in range(len(xls_df)):
                        if type(xls_df["每日行情失敗"][i]) != str :
                            xls_df["每日行情失敗"][i] = stock
                            break
                    if switch_big == 0:  
                        xls_df.to_excel("D:\\Stock Investment\\個股計算開關-市.xlsx")
                    if switch_big == 1: 
                        xls_df.to_excel("D:\\Stock Investment\\個股計算開關-櫃.xlsx")
    if switch == 1:
        xls_df_report.index = range(len(xls_df_report)) 
        xls_df_report.to_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv", encoding = "big5")   
        if int(xls_df_report["日期"][len(xls_df_report)-1]) % 3 == 0 :
            xls_df_report.to_csv("C:\\Users\\User\\Desktop\\Daily Report-Spare\\"+stock+"-day.csv", encoding = "big5")
        print(stock+"每日行情-日 計算完成")
        print(str(datetime.datetime.now()))
    else:
        print(stock+"日報表-每日行情 不需變動")
        print(str(datetime.datetime.now()))                        
    

def week_count_and_tdcc(table,i,stock,switch_big):
    
    switch,switch_,cycle_week = 0,0,11
    data_list = []                
    xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-每日行情.xls")
    xls_df_o = xls_file.parse(xls_file.sheet_names[0], index_col=[0])
                    
    
    try:
        xls_df_week = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-week.csv",engine='python', encoding = "big5",index_col=0)
    except:
        print(stock+"周報表-每日行情 計算失敗")
        print(str(datetime.datetime.now()))   
        if switch_big == 0 :
            xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
            xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0])         
            for i in range(len(xls_df)):
                if type(xls_df["周行情失敗"][i]) != str :
                    xls_df["周行情失敗"][i] = stock
                    break
            xls_df.to_excel("D:\\Stock Investment\\個股計算開關-市.xlsx") 

        elif switch_big == 1 :
            xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-櫃.xlsx")
            xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0])         
            for i in range(len(xls_df)):
                if type(xls_df["周行情失敗"][i]) != str :
                    xls_df["周行情失敗"][i] = stock
                    break
            xls_df.to_excel("D:\\Stock Investment\\個股計算開關-櫃.xlsx") 

                
###防止收盤價有"---"
    for k in range(len(xls_df_o)):
        if type(xls_df_o["收盤價"][k]) != str:
            date_day = int(xls_df_o["日期"][k])                
            break
        
    today = datetime.datetime(int(str(date_day)[:4]),int(str(date_day)[4:6]),int(str(date_day)[6:]))    
    date_week = int(xls_df_week["日期"][len(xls_df_week)-1])
    last_date = datetime.datetime(int(str(date_week)[:4]),int(str(date_week)[4:6]),int(str(date_week)[6:]))

    add_date,need_date = [],[]
    for k in range(len(xls_df_o)):
        if int(xls_df_o["日期"][k]) > date_week :
            if type(xls_df_o["收盤價"][k]) != str:
                add_date.append(int(xls_df_o["日期"][k]))
        else:
            break

    if len(add_date) != 0:
        add_date.sort()
        if len(add_date) > 1:
            for k in range(1,len(add_date)):
                if be_format(str(add_date[k]))[1]-1 == be_format(str(add_date[k-1]))[1]:
                    need_date.append(add_date[k-1])
        
        category = table.row_values(i)[1]
        
        if "工業" in category:
            loca = category.strip().index("工")
            category = category.strip()[:loca]
        elif "業" in category and category != "觀光事業" and category != "農業科技業" and category != "航運業":
            loca = category.strip().index("業")
            category = category.strip()[:loca]
        elif  category == "農業科技業" :
            category = category.strip()[:4] 
        if category == "金融保險" and switch_big == 1 :
            category = "金融"  
                    
        if date_day == date_week:
            print(stock + "-周不須變動")
        elif last_date.weekday() == 4 and today.isocalendar()[1] - last_date.isocalendar()[1] >= 1 :
            switch_ = 1
        elif date_day - date_week == 0 and today.isocalendar()[2] == 5 and last_date.isocalendar()[2] == 5 and len(need_date) == 0:
            print(stock + "-周不須變動")
        elif today.isocalendar()[1] - last_date.isocalendar()[1] == 0  :
            xls_df_week = xls_df_week.drop([len(xls_df_week)-1]) 
            switch_ = 1
    
        if date_day not in need_date:
            need_date.append(date_day)

        if switch_ == 1 :    
            for date_day in need_date: 
                today = datetime.datetime(int(str(date_day)[:4]),int(str(date_day)[4:6]),int(str(date_day)[6:]))
                data_list = []   
                if switch_big == 0 :        
                    xls_file_big_w = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤_周_計算用.xlsx")
                    xls_df_big_w = xls_file_big_w.parse(xls_file_big_w.sheet_names[0], index_col=[0]) 
                elif switch_big == 1 :                    
                    xls_file_big_w = pd.ExcelFile("D:\\Stock Investment\\股票數據\\櫃盤_周_計算用.xlsx")
                    xls_df_big_w = xls_file_big_w.parse(xls_file_big_w.sheet_names[0], index_col=[0]) 
                    
                try:
                    data_list.append([date_day,daily_count_func.HL_price(today,xls_df_o,xls_df_ratio)[3],daily_count_func.HL_price(today,xls_df_o,xls_df_ratio)[5],daily_count_func.HL_price(today,xls_df_o,xls_df_ratio)[4],xls_df_o["開盤價"][xls_df_o["日期"].tolist().index(date_day)],xls_df_o["最高價"][xls_df_o["日期"].tolist().index(date_day)],xls_df_o["最低價"][xls_df_o["日期"].tolist().index(date_day)],xls_df_o["收盤價"][xls_df_o["日期"].tolist().index(date_day)],"","","","","","","",daily_count_func.HL_price(today,xls_df_o,xls_df_ratio)[0],daily_count_func.HL_price(today,xls_df_o,xls_df_ratio)[1],daily_count_func.HL_price(today,xls_df_o,xls_df_ratio)[2],xls_df_o["收盤價"][xls_df_o["日期"].tolist().index(date_day)]])
                    for num in range(len(xls_df_week.columns)-len(data_list[0])):
                        data_list[0].append(0.0001)
        
                    df = pd.DataFrame(data_list,columns= xls_df_week.columns )
                    xls_df_week = pd.concat([xls_df_week,df],axis = 0)
                    xls_df_week.index = range(len(xls_df_week))    
                    
                    xls_df_week["周5日均線"][len(xls_df_week)-1] = daily_count_func.ma_week(xls_df_week,5)
                    xls_df_week["周10日均線"][len(xls_df_week)-1] = daily_count_func.ma_week(xls_df_week,10)
                    xls_df_week["周均線"][len(xls_df_week)-1] = daily_count_func.ma_week(xls_df_week,20) 
                    xls_df_week["周上軌"][len(xls_df_week)-1] = daily_count_func.week_bbands(xls_df_week)[0]
                    xls_df_week["周中軸"][len(xls_df_week)-1] = daily_count_func.week_bbands(xls_df_week)[1]     
                    xls_df_week["周下軌"][len(xls_df_week)-1] = daily_count_func.week_bbands(xls_df_week)[2]           
          
                    for k in range(15):
                        if int(xls_df_ratio["日期"][k]) == date_day:
                            break
                    xls_df_week["參考指標_51"][len(xls_df_week)-1] = xls_df_ratio["周加權總量"][0]
                    xls_df_week["參考指標_52"][len(xls_df_week)-1] = daily_count_func.HL_price(today,xls_df_o,xls_df_ratio)[6]
                    xls_df_week["參考指標_53"][len(xls_df_week)-1] = daily_count_func.HL_price(today,xls_df_o,xls_df_ratio)[8]
                    xls_df_week["參考指標_54"][len(xls_df_week)-1] = daily_count_func.HL_price(today,xls_df_o,xls_df_ratio)[7]
                    xls_df_week["參考指標_55"][len(xls_df_week)-1] = daily_count_func.HL_price(today,xls_df_o,xls_df_ratio)[9]
        
                    week_category_volume = daily_count_func.week_category_volume(xls_df_week,xls_df_big_w,category,cycle_week,switch_big)
                    week_momentum_by_volume = daily_count_func.week_momentum_by_volume(xls_df_week,week_category_volume,cycle_week)
                    week_momentum_by_trans = daily_count_func.week_momentum_by_trans(xls_df_week,week_category_volume,cycle_week)             

                    if category == "金融" and switch_big == 1:
                        category = "金融保險"         
                    elif category == "航運" and switch_big == 1:
                        category = "航運業"  
                            
                    xls_df_week[category+"參考指標_56"][len(xls_df_week)-1] = week_category_volume[4]
                    xls_df_week[category+"參考指標_57"][len(xls_df_week)-1] = week_category_volume[5]
                    xls_df_week[category+"參考指標_58"][len(xls_df_week)-1] = week_category_volume[6]
                    
                    xls_df_week["參考指標_59"][len(xls_df_week)-1] = week_category_volume[0]
                    xls_df_week["參考指標_60"][len(xls_df_week)-1] = week_category_volume[1]
                    xls_df_week["參考指標_61"][len(xls_df_week)-1] = week_category_volume[2]
                    xls_df_week["參考指標_62"][len(xls_df_week)-1] = week_category_volume[3]    
                    
                    xls_df_week["參考指標_63"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_63"][0]
                    xls_df_week["參考指標_64"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_64"][0]
                    xls_df_week["參考指標_65"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_65"][0]
                    xls_df_week["參考指標_66"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_66"][0]
                    xls_df_week["參考指標_67"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_67"][0]
                    xls_df_week["參考指標_68"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_68"][0]       
                    xls_df_week["參考指標_69"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_69"][0]
                    xls_df_week["參考指標_70"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_70"][0]
                    xls_df_week["參考指標_71"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_71"][0]
                    xls_df_week["參考指標_72"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_72"][0]           
                    xls_df_week["參考指標_73"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_73"][0]
                    xls_df_week["參考指標_74"][len(xls_df_week)-1] = week_momentum_by_trans["參考指標_74"][0]    
        
                    xls_df_week["參考指標_75"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_75"][0]
                    xls_df_week["參考指標_76"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_76"][0]
                    xls_df_week["參考指標_77"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_77"][0]
                    xls_df_week["參考指標_78"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_78"][0]
                    xls_df_week["參考指標_79"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_79"][0]
                    xls_df_week["參考指標_80"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_80"][0]
                    xls_df_week["參考指標_81"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_81"][0]
                    xls_df_week["參考指標_82"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_82"][0]
                    xls_df_week["參考指標_83"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_83"][0]
                    xls_df_week["參考指標_84"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_84"][0]           
                    xls_df_week["參考指標_85"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_85"][0]
                    xls_df_week["參考指標_86"][len(xls_df_week)-1] = week_momentum_by_volume["參考指標_86"][0]            
                    switch = 1
                except:
                    print(stock+"周報表-每日行情 計算失敗")
                    print(str(datetime.datetime.now()))   
                    if switch_big == 0 :
                        xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
                        xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0])         
                        for i in range(len(xls_df)):
                            if type(xls_df["周行情失敗"][i]) != str :
                                xls_df["周行情失敗"][i] = stock
                                break
                        xls_df.to_excel("D:\\Stock Investment\\個股計算開關-市.xlsx") 
            
                    elif switch_big == 1 :
                        xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-櫃.xlsx")
                        xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0])         
                        for i in range(len(xls_df)):
                            if type(xls_df["周行情失敗"][i]) != str :
                                xls_df["周行情失敗"][i] = stock
                                break
                        xls_df.to_excel("D:\\Stock Investment\\個股計算開關-櫃.xlsx") 
                       
            if switch == 1:                    
                xls_df_week.to_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-week.csv", encoding = "big5")  
                if int(xls_df_week["日期"][len(xls_df_week)-1]) % 2 == 0 :
                    xls_df_week.to_csv("C:\\Users\\User\\Desktop\\Daily Report-Spare\\"+stock+"-week.csv", encoding = "big5")
        
                print(stock+"周報表-每日行情 計算完成")
                print(str(datetime.datetime.now()))    

###################################
# 主要運行程式碼-每日報表_新-力道計算 #
#################################  
def daily_power_count(stock,category,switch_big):
    
    cycle_week,cycle_day,switch = 11,59,0
    
    if "工業" in category:
        loca = category.strip().index("工")
        category = category.strip()[:loca]
    elif "業" in category and category != "觀光事業" and category != "航運業" and switch_big == 0:
        loca = category.strip().index("業")
        category = category.strip()[:loca]
    elif "業" in category and category != "觀光事業" and category != "農業科技業" and switch_big == 1:
        loca = category.strip().index("業")
        category = category.strip()[:loca]
    elif  category == "農業科技業" and switch_big == 1:
        category = category.strip()[:4] 
    if category == "金融保險" and switch_big == 1:
        category = "金融"   
      
    xls_df_report = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)
    week_table = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-week.csv",engine='python', encoding = "big5",index_col=0)
   
    date_list = []    
    for j in range(1,len(xls_df_report["日期"])):        
        if float(xls_df_report["參考指標"][len(xls_df_report["日期"])-j]) == 0.0001:
            date_list.append(xls_df_report["日期"][len(xls_df_report["日期"])-j])
        else:
            break

    if len(date_list) == 0:
        print(stock+"日報表-力道計算 不需變動")
        print(str(datetime.datetime.now()))   
        return None
 
    if switch_big == 0:    
        xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤_計算用.xlsx")
        xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])    
        xls_df_big = erase_for_big(xls_file_big_d,xls_df_report)

        xls_file_big_w = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤_周_計算用.xlsx")
        xls_df_big_w = xls_file_big_w.parse(xls_file_big_w.sheet_names[0], index_col=[0]) 
        xls_df_big_w = erase_for_big_w(xls_df_big_w,week_table)    
    
    elif switch_big == 1:
        xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\櫃盤_計算用.xlsx")
        xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])    
        xls_df_big = erase_for_big(xls_file_big_d,xls_df_report)
            
        xls_file_big_w = pd.ExcelFile("D:\\Stock Investment\\股票數據\\櫃盤_周_計算用.xlsx")
        xls_df_big_w = xls_file_big_w.parse(xls_file_big_w.sheet_names[0], index_col=[0])
        xls_df_big_w = erase_for_big_w(xls_df_big_w,week_table)

    date_list.sort()
    for data_date in date_list  :                         
        for k in range(len(xls_df_report["日期"])-60,len(xls_df_report["日期"])):
            if int(xls_df_report["日期"][k]) == int(data_date):
                break

        if category == "金融保險" and switch_big == 1:
            category = "金融"  
        elif category == "航運業" and switch_big == 1:
            category = "航運"

        try:
            date_ = data_date                    
            day_category_trans = daily_count_func.day_category_trans(xls_df_report,xls_df_big,category,cycle_day,k)
            day_momentum_by_trans = daily_count_func.day_momentum_by_trans(xls_df_report,day_category_trans,week_table,xls_df_big,category,xls_df_big_w,cycle_day,cycle_week,k,date_)
            day_category_volume = daily_count_func.day_category_volume(xls_df_report,xls_df_big,category,cycle_day,k)
            day_momentum_by_volume = daily_count_func.day_momentum_by_volume(xls_df_report,day_category_volume,week_table,xls_df_big,category,xls_df_big_w,cycle_day,cycle_week,date_,k)
            
            xls_df_report["參考指標_1"][k] = day_momentum_by_trans["參考指標_1"][0]
            xls_df_report["參考指標_2"][k] = day_momentum_by_trans["參考指標_2"][0]
            xls_df_report["參考指標_3"][k] = day_momentum_by_trans["參考指標_3"][0]
            xls_df_report["參考指標_4"][k] = day_momentum_by_trans["參考指標_4"][0]
            xls_df_report["參考指標_5"][k] = day_momentum_by_trans["參考指標_5"][0]
            xls_df_report["參考指標_6"][k] = day_momentum_by_trans["參考指標_6"][0]
            xls_df_report["參考指標_7"][k] = day_momentum_by_trans["參考指標_7"][0]
            xls_df_report["參考指標_8"][k] = day_momentum_by_trans["參考指標_8"][0]
            xls_df_report["參考指標_9"][k] = day_momentum_by_trans["參考指標_9"][0]
            xls_df_report["參考指標_10"][k] = day_momentum_by_trans["參考指標_10"][0]
            xls_df_report["參考指標_11"][k] = day_category_trans[0]
            xls_df_report["參考指標_12"][k] = day_category_trans[1]
            xls_df_report["參考指標_13"][k] = day_category_trans[2]
            if category == "金融" and switch_big == 1:
                category = "金融保險"         
            elif category == "航運" and switch_big == 1:
                category = "航運業"             
            xls_df_report[category + "參考指標"][k] = day_category_trans[3]
            xls_df_report["參考指標_14"][k] = day_momentum_by_trans["參考指標_14"][0]    
            xls_df_report["參考指標_15"][k] = day_momentum_by_trans["參考指標_15"][0]
            xls_df_report["參考指標_16"][k] = day_momentum_by_trans["參考指標_16"][0]
            xls_df_report["參考指標_17"][k] = day_momentum_by_trans["參考指標_17"][0]
            xls_df_report["參考指標_18"][k] = day_momentum_by_trans["參考指標_18"][0]
                
            xls_df_report["參考指標_19"][k] = day_category_volume[0]      
            xls_df_report["參考指標_20"][k] = day_category_volume[1]        
            xls_df_report["參考指標_21"][k] = day_category_volume[2]  
            xls_df_report[category+"參考指標_22"][k] = day_category_volume[5]              
            xls_df_report["參考指標_23"][k] = daily_count_func.day_category_stock(xls_df_report,xls_df_big,category,cycle_day,k,switch_big)[0] 
            xls_df_report["參考指標_24"][k] = daily_count_func.day_category_stock(xls_df_report,xls_df_big,category,cycle_day,k,switch_big)[1] 
            xls_df_report["參考指標_25"][k] = daily_count_func.day_category_stock(xls_df_report,xls_df_big,category,cycle_day,k,switch_big)[2] 
            xls_df_report["參考指標_26"][k] = daily_count_func.day_category_money(xls_df_report,xls_df_big,category,cycle_day,k,switch_big)[0] 
            xls_df_report["參考指標_27"][k] = daily_count_func.day_category_money(xls_df_report,xls_df_big,category,cycle_day,k,switch_big)[1] 
            xls_df_report["參考指標_28"][k] = daily_count_func.day_category_money(xls_df_report,xls_df_big,category,cycle_day,k,switch_big)[2]             
            xls_df_report[category+"參考指標_29"][k] = day_category_volume[4]                                 
            xls_df_report["參考指標_30"][k] = day_momentum_by_volume["參考指標_30"][0]
            xls_df_report["參考指標_31"][k] = day_momentum_by_volume["參考指標_31"][0]
            xls_df_report["參考指標_32"][k] = day_momentum_by_volume["參考指標_32"][0]
            xls_df_report["參考指標_33"][k] = day_momentum_by_volume["參考指標_33"][0]
            xls_df_report["參考指標_34"][k] = day_momentum_by_volume["參考指標_34"][0]
            xls_df_report["參考指標_35"][k] = day_momentum_by_volume["參考指標_35"][0]
            xls_df_report["參考指標_36"][k] = day_momentum_by_volume["參考指標_36"][0]
            xls_df_report["參考指標_37"][k] = day_momentum_by_volume["參考指標_37"][0]
            xls_df_report["參考指標_38"][k] = day_momentum_by_volume["參考指標_38"][0]
            xls_df_report["參考指標_39"][k] = day_momentum_by_volume["參考指標_39"][0]                  
            xls_df_report["參考指標_40"][k] = day_momentum_by_volume["參考指標_40"][0]
            xls_df_report["參考指標_41"][k] = day_momentum_by_volume["參考指標_41"][0]
            xls_df_report["參考指標_42"][k] = day_momentum_by_volume["參考指標_42"][0]
            xls_df_report["參考指標_43"][k] = day_momentum_by_volume["參考指標_43"][0]
            xls_df_report["參考指標_44"][k] = day_momentum_by_volume["參考指標_44"][0]
            xls_df_report["參考指標_45"][k] = day_momentum_by_volume["參考指標_45"][0]    
                    
            switch = 1
        except:
            print(stock+"日報表-力道計算失敗")
            print(str(datetime.datetime.now())) 
            if switch_big == 0:   
                xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-市.xlsx")
            if switch_big == 1:   
                xls_file = pd.ExcelFile("D:\\Stock Investment\\個股計算開關-櫃.xlsx")                    
            xls_df = xls_file.parse(xls_file.sheet_names[0], index_col=[0])         
            for i in range(len(xls_df)):
                if type(xls_df["力道計算失敗"][i]) != str :
                    xls_df["力道計算失敗"][i] = stock
                    break
            if switch_big == 0:  
                xls_df.to_excel("D:\\Stock Investment\\個股計算開關-市.xlsx")
            if switch_big == 1: 
                xls_df.to_excel("D:\\Stock Investment\\個股計算開關-櫃.xlsx")                                        

    if switch == 1:
        xls_df_report.index = range(len(xls_df_report)) 
        xls_df_report.to_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv", encoding = "big5") 
        if int(xls_df_report["日期"][len(xls_df_report)-1]) % 3 == 0 :
            xls_df_report.to_csv("C:\\Users\\User\\Desktop\\Daily Report-Spare\\"+stock+"-day.csv", encoding = "big5")        
        print(stock+"日報表-力道計算完成")
        print(str(datetime.datetime.now()))
    else:
        print(stock+"日報表-力道計算不需變動")
        print(str(datetime.datetime.now()))                    
   
#################################    
### 主要運行程式碼-每日報表_計算 #
###############################       
def daily_category_count(stock,xls_df_basic,no_need,switch_big,category): 
    
    if stock != '' and stock not in no_need: 
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
                
        xls_df_day = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)
        xls_df_week = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-week.csv",engine='python', encoding = "big5",index_col=0)   

        if switch_big == 0 :                    
            xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤_計算用.xlsx")
            xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])    
            xls_df_big = erase_for_big(xls_file_big_d,xls_df_day) 

        elif switch_big == 1 :     
            xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\櫃盤_計算用.xlsx")
            xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])    
            xls_df_big = erase_for_big(xls_file_big_d,xls_df_day) 

        locate = excel_locate(xls_df_basic,stock)
        
        if int(xls_df_day["日期"][len(xls_df_day)-1]) == int(xls_df_week["日期"][len(xls_df_week)-1]) :
            price_slope_daily_new.price_range_day(stock,xls_df_day,xls_df_basic,locate)
            price_slope_daily_new.perb_data_show(stock,xls_df_day,xls_df_basic,locate)
            price_slope_daily_new.meet_20days_standard_within_60days(stock,xls_df_day,xls_df_big,xls_df_basic,locate)
            price_slope_daily_new.meet_20days_standard_today(stock,xls_df_day,xls_df_big,xls_df_basic,locate)
            price_slope_daily_new.up_down_day(stock,xls_df_day,xls_df_basic,locate)
            price_slope_daily_new.volume_day(stock,xls_df_day,xls_df_basic,locate)
            price_slope_daily_new.signal_wait_for_5ma(stock,xls_df_day,xls_df_basic,locate)
            xls_df_basic["收盤價"][locate] = xls_df_day["日收盤價"][len(xls_df_day["日收盤價"])-1]
            print(stock+"已經完成計算")
            print(str(datetime.datetime.now()))                              
            
