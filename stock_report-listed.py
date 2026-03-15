# -*- coding: utf-8 -*-
"""
Created on Wed Mar 14 23:41:32 2018

@author: YAO
"""
import xlrd
import pandas as pd
import index_all
import numpy as np
from datetime import datetime

###個股若交易不活絡，該周會有無收盤價情形，因此刪除
def week_format():
    for j in range(len(week_table)):
        if week_table["周收盤價"][j] == "" :
            week_table.drop([j],axis = 0, inplace = True)
    week_table.index = range(len(week_table)) 
    return week_table

###個股若交易不活絡，該日會有無收盤價情形，因此刪除
def erase_data(xls_df_o):
    global erase_locate
    erase_locate = []
    column_title_new =['日期',"成交股數","成交筆數","成交金額","日開盤價","日最高價","日最低價","日收盤價"] 
    xls_df_o = xls_df_o.drop("最後揭示買價", axis = 1)
    xls_df_o = xls_df_o.drop("最後揭示買量", axis = 1)
    xls_df_o = xls_df_o.drop("最後揭示賣價", axis = 1)
    xls_df_o = xls_df_o.drop("最後揭示賣量", axis = 1)
    xls_df_o = xls_df_o.drop("本益比", axis = 1)   
    xls_df_o = xls_df_o.drop("漲跌價差", axis = 1)
    for i in range(len(xls_df_o)):
        try:
            if xls_df_o["收盤價"][i].strip() == "--" or xls_df_o["收盤價"][i].strip() == "----":
                xls_df_o.drop([i],axis = 0, inplace = True)
                erase_locate.append(i)
        except:
            continue
    be_array = np.array(xls_df_o)  #將表格用成array格式在合併，將index重新排列
    normal_data = pd.DataFrame(be_array,columns = column_title_new) 
    return normal_data

###個股不一定每個交易日皆有交易，因此大盤指數的交易資料日期要與個股的交易資訊日期一樣
def erase_for_big(xls_df_big,xls_df):
    xls_df_date = set(list(xls_df["日期"]))
    xls_df_big_date = set(list(xls_df_big["日期"]))
    diff = xls_df_big_date.difference(xls_df_date)
    diff = list(diff)   
    diff.sort(reverse = False)
    
    for i in range(len(xls_df_big)):
        if xls_df_big["日期"][i] in diff :
            xls_df_big.drop([i],axis = 0, inplace = True)
    xls_df_big.index = range(len(xls_df_big))
    
    return xls_df_big

###可以選擇報表日期長度起訖日
def cut_file(file):
    file = file.sort_values(by = "日期",ascending = False)
    file.index = range(len(file))
    for k in range(len(file)):
        if int(file["日期"][k]) <= 20230928:
            break
    for p in range(len(file)):
        if int(file["日期"][p]) <= 20211004:
            break

    file = file[k:p]
    file.index = range(len(file))
            
    return file        
            
file = xlrd.open_workbook(r"D:\\Stock Investment\\股票分類-下載用(去除重複)-上市.xlsx")
table = file.sheets()[0]
sheet = file.sheets()[1]

fail_list,fail_list1,cycle_day,cycle_week = [],[],60,12

for i in range(33):
    stock_list = table.row_values(i)[2:]
    number_list = sheet.row_values(i)[2:]
    for stock in stock_list:
        data_list = []  
        if stock != "" : 
            try:
                category = table.row_values(i)[1]
                number = str(int(number_list[stock_list.index(stock)]))
                xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\"+table.row_values(i)[1]+"\\"+stock+"\\"+stock+"-每日行情.xls")
                xls_df_o = xls_file.parse(xls_file.sheet_names[0], index_col=[0])
                xls_df_o = erase_data(xls_df_o)
                xls_df_o = index_all.volume_adjustment(xls_df_o,table,stock,i)
                xls_df_o = index_all.day(xls_df_o)
                xls_df = cut_file(xls_df_o)
                    
                xls_file_big = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤.xlsx")
                xls_df_big = xls_file_big.parse(xls_file_big.sheet_names[0], index_col=[0])
                xls_df_big = erase_for_big(xls_df_big,xls_df)
                xls_df_big = xls_df_big.sort_values(by = "日期",ascending = True)
                xls_df_big.index = range(len(xls_df_big))
    
                xls_file_big_w = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤_周.xlsx")
                xls_df_big_w = xls_file_big_w.parse(xls_file_big_w.sheet_names[0], index_col=[0])
                xls_df_big_w = (xls_df_big_w,xls_df) 
                xls_df_big_w = xls_df_big_w.sort_values(by = "日期",ascending = True)
                xls_df_big_w.index = range(len(xls_df_big_w))
                
                normal_file = index_all.normal(xls_df)        
###week        
                week_information = index_all.week_k(normal_file)
                week_k = week_information[0]
                week_data_pure = week_information[1]
                week_close_price = week_information[2]
                week_locate = week_information[3]
                week_float = week_information[4]
                week_moving_average = index_all.week_moving_average(normal_file,week_float,week_locate,week_close_price)
                week_ma20_ma5_bias = index_all.week_ma20_ma5_bias(normal_file,week_float,week_locate)
                week_macd = index_all.week_macd(normal_file,week_float,week_locate)
                week_kd = index_all.week_kd(normal_file,week_k,week_locate,number)
                week_bbands_information = index_all.week_bbands(normal_file,week_k,week_locate,week_moving_average)
                week_bbands = week_bbands_information[0]
                week_table = pd.concat([normal_file,week_moving_average,week_bbands],axis = 1)  
                week_table = index_all.week_volume_money_sum(week_table)
                week_category_volume = index_all.week_category_volume(week_table,xls_df_big_w,category,cycle_week)
                week_momentum_by_volume = index_all.week_momentum_by_volume(week_table,week_category_volume,cycle_week)
                week_momentum_by_trans = index_all.week_momentum_by_trans(week_table,week_category_volume,cycle_week)
                week_table = pd.concat([week_table,week_momentum_by_volume,week_category_volume,week_momentum_by_trans],axis = 1) 
##day                         

                day_category_money = index_all.day_category_money(normal_file,xls_df_big,category,cycle_day)
                day_category_stock = index_all.day_category_stock(normal_file,xls_df_big,category,cycle_day)
                day_moving_average = index_all.day_moving_average(normal_file)
                day_moving_volume = index_all.moving_average_volume(normal_file)
                day_ma = index_all.day_ma(day_moving_average)
                day_kd = index_all.day_kd(normal_file,number)
                day_macd = index_all.day_macd(normal_file)                   
                day_ma20_ma60_bias = index_all.day_ma20_ma60_bias(normal_file,day_moving_average)
                day_bbands = index_all.day_bbands(normal_file,day_moving_average,cycle_day,number)
                day_perb = index_all.day_perb(normal_file,day_bbands)
                day_band_width = index_all.day_band_width(day_bbands)
                day_bbands = index_all.day_bbands(day_bbands)
                day_pattern = index_all.pattern(normal_file)
                day_down_up_percentage = index_all.day_down_up_percentage(normal_file)
                category_money_tracking = index_all.category_money_tracking(normal_file,erase_locate,cycle_day)
                day_category_volume = index_all.day_category_volume(normal_file,xls_df_big,category,cycle_day,category_money_tracking,day_category_stock,day_category_money,erase_locate)   
                day_momentum_by_volume = index_all.day_momentum_by_volume(normal_file,day_category_volume,week_table,xls_df_big,category,xls_df_big_w,week_category_volume,cycle_day,cycle_week)                           
                day_category_trans =  index_all.day_category_trans(normal_file,xls_df_big,category,cycle_day)             
                day_momentum_by_trans = index_all.day_momentum_by_trans(normal_file,day_category_trans,week_table,xls_df_big,category,xls_df_big_w,week_category_volume,cycle_day,cycle_week)                        
                new_table = pd.concat([normal_file,day_down_up_percentage,day_moving_average,day_kd,day_macd,day_ma20_ma60_bias,day_moving_volume,day_ma,day_category_money,day_category_stock,day_momentum_by_volume,day_momentum_by_trans,day_category_trans,day_bbands,day_bbands,day_category_volume,day_ma20_ma60_bias,day_perb,day_band_width,category_money_tracking,day_pattern],axis = 1)       

            except:
                 fail_list.append(stock)
                 print(fail_list)
                 continue
###匯出檔案

            new_table.to_csv(r"C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv", encoding = "big5")
            week_table.to_csv(r"C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-week.csv", encoding = "big5")
            print(stock+"已經下載")
            print(str(datetime.now()))
