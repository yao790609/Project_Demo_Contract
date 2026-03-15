# -*- codinD: utf-8 -*-
"""
Created on Thu Apr 26 13:50:11 2018

@author: YAO
"""

import pandas as pd
import math

def need_days_pair(file,days,list_need):
            
    for j in range(len(list_need)):
        locals()['xls_file_%s' % j] = 0
        for k in range(1,30):
            if file[list_need[j]][len(file)-k] == 100 :
                locals()['xls_file_%s' % j] = locals()['xls_file_%s' % j] + 1
            elif file[list_need[j]][len(file)-k] > file[list_need[j]][len(file)-k-1] :
                locals()['xls_file_%s' % j] = locals()['xls_file_%s' % j] + 1
            else:
                locals()['xls_file_%s' % j] = locals()['xls_file_%s' % j] + 1
                n = k
                break
        
        if k == 29 :
            n = 29
            
        for k in range(1+n,60):                     
            if  locals()['xls_file_%s' % j] < days:
                locals()['xls_file_%s' % j] = locals()['xls_file_%s' % j] + 1
            elif float(file[list_need[j]][len(file)-k]) > float(file[list_need[j]][len(file)-k-1]):
                locals()['xls_file_%s' % j] = locals()['xls_file_%s' % j] + 1
            else:
                break
            
    if locals()['xls_file_%s' % 0] > locals()['xls_file_%s' % 1]:
        need_days = locals()['xls_file_%s' % 0]
    else:
        need_days = locals()['xls_file_%s' % 1]
    
    return need_days

def need_days(file,days,list_need):
            
    for j in range(len(list_need)):
        locals()['xls_file_%s' % j] = 0
        for k in range(1,30):
            if file[list_need[j]][len(file)-k] == 100 :
                locals()['xls_file_%s' % j] = locals()['xls_file_%s' % j] + 1
            elif file[list_need[j]][len(file)-k] > file[list_need[j]][len(file)-k-1] :
                locals()['xls_file_%s' % j] = locals()['xls_file_%s' % j] + 1
            else:
                locals()['xls_file_%s' % j] = locals()['xls_file_%s' % j] + 1
                n = k
                break
        
        if k == 29 :
            n = 29
            
        for k in range(1+n,60):                     
            if  locals()['xls_file_%s' % j] < days:
                locals()['xls_file_%s' % j] = locals()['xls_file_%s' % j] + 1
            elif float(file[list_need[j]][len(file)-k]) > float(file[list_need[j]][len(file)-k-1]):
                locals()['xls_file_%s' % j] = locals()['xls_file_%s' % j] + 1
            else:
                break

    return locals()['xls_file_%s' % 0]

def volume_day(stock,xls_df_day,xls_df_basic,locate):
    
    if "," in str(xls_df_day["5日均量"][len(xls_df_day["5日均量"])-1]):
        xls_df_basic["5日均量"][locate] = str(int(int(xls_df_day["5日均量"][len(xls_df_day["5日均量"])-1].replace(",",""))/1000))
    else:        
        xls_df_basic["5日均量"][locate] = str(int(int(xls_df_day["5日均量"][len(xls_df_day["5日均量"])-1])/1000))
        
def up_down_day(stock,xls_df_day,xls_df_basic,locate):
    
    total = 0
    for i in range(3):
        total = total + xls_df_day["漲跌幅"][len(xls_df_day)-1-i]
        
    xls_df_basic["3日漲跌幅"][locate] = total        
    xls_df_basic["漲跌幅"][locate] = xls_df_day["漲跌幅"][len(xls_df_day)-1]

def perb_data_show(stock,xls_df_day,xls_df_basic,locate):
   
    xls_df_basic["日價格日布林 %B"][locate] = round((xls_df_day["日收盤價"][len(xls_df_day ["日收盤價"])-1]-xls_df_day["日下軌"][len(xls_df_day ["日下軌"])-1])/(xls_df_day["日上軌"][len(xls_df_day ["日上軌"])-1]-xls_df_day["日下軌"][len(xls_df_day ["日上軌"])-1])*100,2) 
    
def day_price_week_bollin(stock,xls_df_day,xls_df_basic,locate,xls_df_week):    
    
    close_price_day = xls_df_day["日收盤價"][len(xls_df_day["日收盤價"])-1]
    i =  len(xls_df_week ["周收盤價"])-1      
    b_value = round((float(close_price_day)-xls_df_week["周下軌"][i])/(xls_df_week["周上軌"][i]-xls_df_week["周下軌"][i])*100,2)
    
    xls_df_basic["日價格周布林 %B"][locate] = str(b_value)
    
def price_range_day(stock,xls_df_day,xls_df_basic,locate):
    close_price = xls_df_day["日收盤價"]
    ma20_day = xls_df_day["20日均線"]
    ma5_day = xls_df_day["5日均線"]

    i =  len(xls_df_day["日收盤價"])-1       
    if close_price[i] >= ma20_day[i]:
        if close_price[i-1] >= ma5_day[i-1]:
            if close_price[i] < ma5_day[i]:
                xls_df_basic["日價格情形"][locate] = "日5均>=日價格>=日20均"
            else:            
                xls_df_basic["日價格情形"][locate] = "日價格>=日5均>=日20均"
        elif close_price[i] > ma5_day[i]:
            xls_df_basic["日價格情形"][locate] = "日價格>=日5均>=日20均"
        elif close_price[i] <= ma5_day[i]:
            xls_df_basic["日價格情形"][locate] = "日5均>=日價格>=日20均"
                
    elif close_price[i] <= ma20_day[i]:
        if close_price[i-1] <= ma5_day[i-1]:
            if close_price[i] > ma5_day[i]:
                xls_df_basic["日價格情形"][locate] = "日20均>=日價格>=日5均"  
            else:
                xls_df_basic["日價格情形"][locate] = "日20均>=日5均>=日價格"                
        elif close_price[i] <= ma5_day[i]:
            xls_df_basic["日價格情形"][locate] = "日20均>=日5均>=日價格"
        elif close_price[i] > ma5_day[i]:
            xls_df_basic["日價格情形"][locate] = "日20均>=日價格>=日5均"
                      
def price_range_week(stock,xls_df_day,xls_df_basic,locate,xls_df_week):
    close_price = xls_df_week["周收盤價"]
    close_price_day = xls_df_day["日收盤價"][len(xls_df_day)-1]
    ma5_week = xls_df_week["周5日均線"]          
    ma20_week = xls_df_week["周20日均線"] 

    i =  len(xls_df_week ["周收盤價"])-1          
    if float(close_price[i]) >= float(ma20_week[i]):
        if float(close_price_day) >= float(close_price[i]) >= float(ma5_week[i]):
            xls_df_basic["周價格情形"][locate] = "日價格>=上周價格>=周5均>=周20均"
        elif float(close_price[i]) >= float(close_price_day) > float(ma5_week[i]):
            xls_df_basic["周價格情形"][locate] = "上周價格>=日價格>=周5均>=周20均"
        elif float(close_price[i]) >=float(ma5_week[i]) >= float(close_price_day):
            xls_df_basic["周價格情形"][locate] = "上周價格>=周5均>=日價格>=周20均"
        elif float(close_price_day) >=float(ma5_week[i]) >= float(close_price[i]):
            xls_df_basic["周價格情形"][locate] = "日價格>=周5均>=上周價格>=周20均"   
        elif float(ma5_week[i]) >=float(close_price_day) >= float(close_price[i]):
            xls_df_basic["周價格情形"][locate] = "周5均>=日價格>=上周價格>=周20均"   
        elif float(ma5_week[i]) >= float(close_price[i]) >=float(close_price_day):
            xls_df_basic["周價格情形"][locate] = "周5均>=上周價格>=日價格>=周20均"   
            
    elif float(close_price[i]) <= ma20_week[i]:
        if float(close_price_day) >= float(close_price[i]) >= float(ma5_week[i]):
            xls_df_basic["周價格情形"][locate] = "周20均>=日價格>=上周價格>=周5均"
        elif float(close_price[i]) >= float(close_price_day) >float(ma5_week[i]):
            xls_df_basic["周價格情形"][locate] = "周20均>=上周價格>=日價格>>=周5均"
        elif float(close_price[i]) >= float(ma5_week[i]) >= float(close_price_day):
            xls_df_basic["周價格情形"][locate] = "周20均>=上周價格>=周5均>=日價格"   
        elif float(ma5_week[i])>= float(close_price[i]) >= float(close_price_day):
            xls_df_basic["周價格情形"][locate] = "周20均>=周5均>=上周價格>=日價格"   
        elif float(ma5_week[i]) >= float(close_price_day) >= float(close_price[i]):
            xls_df_basic["周價格情形"][locate] = "周20均>=周5均>=日價格>=上周價格"   
        elif float(close_price_day) >= float(ma5_week[i]) >= float(close_price[i]):
            xls_df_basic["周價格情形"][locate] = "周20均>=日價格>=周5均>=上周價格"   

def meet_20days_standard_within_60days(stock,xls_df_day,xls_df_big,xls_df_basic,locate):

    switch = 0
    for i in range(len(xls_df_day)-60,len(xls_df_day)-1):
        if xls_df_day["日收盤價"][len(xls_df_day)-1] > xls_df_day["日上軌"][len(xls_df_day)-1] and xls_df_day["日收盤價"][len(xls_df_day)-1] > xls_df_day["60日均線"][len(xls_df_day)-1]:            
            switch = 1

        if switch == 1:
            if xls_df_day["日收盤價"][i] > xls_df_day["日上軌"][i] and xls_df_day["日收盤價"][i] > xls_df_day["60日均線"][i] :
                if xls_df_day["動能指標"][i] <= 3 and xls_df_day["動能指標_1"][i] <= 3  :
                    if xls_df_day["動能指標_2"][i] <= 3 and xls_df_day["動能指標_3"][i] <= 3 : 
                        xls_df_basic["訊號-60日內"][locate] = "V"
                                     
def meet_20days_standard_today(stock,xls_df_day,xls_df_big,xls_df_basic,locate):            

    if xls_df_day["日收盤價"][len(xls_df_day)-1] > xls_df_day["日上軌"][len(xls_df_day)-1] and xls_df_day["日收盤價"][len(xls_df_day)-1] > xls_df_day["60日均線"][len(xls_df_day)-1] :
        if xls_df_day["動能指標"][len(xls_df_day)-1] <= 3 and xls_df_day["動能指標_1"][len(xls_df_day)-1] <= 3 :
            if xls_df_day["動能指標_2"][len(xls_df_day)-1] <= 3 and xls_df_day["動能指標_3"][len(xls_df_day)-1] <= 3  : 
                xls_df_basic["訊號-今日"][locate] = "V"
                    
def signal_wait_for_5ma(stock,xls_df_day,xls_df_basic,locate):
    
    switch,time1,time2 = 0,0,0
    for i in range(2,10):                          
        if xls_df_day["日收盤價"][len(xls_df_day)-i] > xls_df_day["日上軌"][len(xls_df_day)-i] and xls_df_day["日收盤價"][len(xls_df_day)-i] > xls_df_day["60日均線"][len(xls_df_day)-i] :
            if xls_df_day["動能指標"][len(xls_df_day)-i] <= 3 and xls_df_day["動能指標_1"][len(xls_df_day)-i] <= 3 :
                if xls_df_day["動能指標_2"][len(xls_df_day)-i] <= 3 and xls_df_day["動能指標_3"][len(xls_df_day)-i] <= 3  : 
                    switch = 1
                    break

    if switch == 1 :
        for j in range(len(xls_df_day)-i,len(xls_df_day)):
            if xls_df_day["日收盤價"][j] < xls_df_day["5日均線"][j]:
                switch,time1 = 0,0
                break
            else:
                time1 = time1 + 1
                
        close_price_theday = xls_df_day["日收盤價"][len(xls_df_day)-i]    
        close_price_today = xls_df_day["日收盤價"][len(xls_df_day)-1]
        ma5_today = xls_df_day["5日均線"][len(xls_df_day)-1]
        
    if switch == 1 :
        message1 = str(time1)+"天,訊號價距離:"+str(round(((close_price_today-close_price_theday)/close_price_theday)*100,2))+"均線:"+str(round(((close_price_today-ma5_today)/ma5_today)*100,2))
        xls_df_basic["訊號發動盤堅天數"][locate] = message1  
        xls_df_basic["訊號價"][locate] = close_price_theday
        
    switch = 0
    for i in range(1,10):
        if xls_df_day["日20角度5均"][len(xls_df_day)-1-i] > xls_df_day["日20角度10均"][len(xls_df_day)-1-i] and xls_df_day["日收盤價"][len(xls_df_day)-1-i] > xls_df_day["日上軌"][len(xls_df_day)-1-i] and xls_df_day["日收盤價"][len(xls_df_day)-1-i] > xls_df_day["60日均線"][len(xls_df_day)-1-i]:              
            switch = 1
            break

    if switch == 1 :
        for j in range(len(xls_df_day)-i,len(xls_df_day)):
            if xls_df_day["日收盤價"][j] < xls_df_day["5日均線"][j]:
                switch,time2 = 0,0      
                break
            else:
                time2 = time2 + 1
                
        close_price_theday2 = xls_df_day["日收盤價"][len(xls_df_day)-i]    
        close_price_today2 = xls_df_day["日收盤價"][len(xls_df_day)-1]
        ma5_today2 = xls_df_day["5日均線"][len(xls_df_day)-1]
        
    switch_1 = 0
    if switch == 1:
        for i in range(len(xls_df_day)-60,len(xls_df_day)-1):
            if xls_df_day["日收盤價"][i] > xls_df_day["日上軌"][i] and xls_df_day["日收盤價"][i] > xls_df_day["60日均線"][i] :
                if xls_df_day["動能指標"][i] <= 3 and xls_df_day["動能指標_1"][i] <= 3  :
                    if xls_df_day["動能指標_2"][i] <= 3 and xls_df_day["動能指標_3"][i] <= 3 : 
                        switch_1 = 1
                                    
    if switch_1 == 1:
        message2 = str(time2)+"天,訊號價距離:"+str(round(((close_price_today2-close_price_theday2)/close_price_theday2)*100,2))+",均線:"+str(round(((close_price_today2-ma5_today2)/ma5_today2)*100,2))
        xls_df_basic["訊號發動盤堅天數"][locate] = message2
        xls_df_basic["訊號價"][locate] = close_price_theday2
        
    if time1 != 0 and time2 != 0 :        
        if time1 > time2:        
            xls_df_basic["訊號發動盤堅天數"][locate] = message1 
            xls_df_basic["訊號價"][locate] = close_price_theday
        else:
            xls_df_basic["訊號發動盤堅天數"][locate] = message2
            xls_df_basic["訊號價"][locate] = close_price_theday2
            
def industry_stock(stock,xls_df_day,xls_df_basic,xls_df_cate,table_industry,table_category,locate,category,xls_df_big):

    switch = 0
    if type(xls_df_basic["訊號發動盤堅天數"][locate]) == str and xls_df_basic["訊號發動盤堅天數"][locate] != "":
        switch = 1
    elif xls_df_basic["訊號-今日"][locate] == "V" or  xls_df_basic["訊號-60日內"][locate] == "V":
        switch = 1

    if switch == 1:
        for j in range(len(xls_df_cate)):
            if xls_df_cate["股票名稱"][j] == stock :
                try:
                    if math.isnan(xls_df_cate["分類"][j]) == True:
                        return None
                except:
                    weight_list = list(set(xls_df_cate["分類"][j].split(",")))                    
                break

        xls_file = pd.ExcelFile("D:\Stock Investment\指標-2022.xlsx")
        xls_df_AD = xls_file.parse(xls_file.sheet_names[0], index_col=[0])

        for k in range(len(xls_df_AD)):
            if int(xls_df_AD["日期"][k]) == int(xls_df_day["日期"][len(xls_df_day)-1]):
                xls_df_AD = xls_df_AD[k:k+100]
                xls_df_AD.index = range(len(xls_df_AD))
                if len(xls_df_AD) >= 60 :
                    switch = 1
                break
     
        xls_df_date = set(list(xls_df_AD["日期"]))
        xls_df_big_date = set(list(xls_df_day["日期"]))
        diff = xls_df_date.difference(xls_df_big_date)
        diff = list(diff)   
        if len(diff) != 0 : 
            diff.sort(reverse = False)
            for i in range(len(xls_df_AD)):
                if xls_df_AD["日期"][i] in diff :
                    xls_df_AD.drop([i],axis = 0, inplace = True)
        xls_df_AD = xls_df_AD.sort_values(by = "日期",ascending = True)        
        xls_df_AD.index = range(len(xls_df_AD))
            
        change_name = {}
        for k in range(len(table_category)):
            change_name.setdefault(table_category["類名稱"][k],table_category["名稱"][k])
            
        cat_uni_perb,cat_price_perb,cat_uni_big,cat_price_big = {},{},{},{}  
        cat_uni_perb_result,cat_price_perb_result,cat_uni_big_result,cat_price_big_result = {},{},{},{}             
        cat_growth_perb,cat_growth_perb_result,cat_growth_big_perb,cat_growth_big_perb_result = {},{},{},{}                      
        stock_cat_weighted,stock_cat_pure,stock_cat_stock,stock_stock_pure = {},{},{},{}  
        cat_stock_pure_,cat_money_pure_,cat_money_,cat_stock_ = {},{},{},{}
        cat_stock_pure_result,cat_money_pure_result,cat_money_result,cat_stock_result = {},{},{},{}  
        stock_cat_weighted_result,stock_cat_pure_result,stock_cat_stock_result,stock_stock_pure_result = {},{},{},{}  
        cat_stock_growth,cat_stock_growth_ratio,cat_stock_growth_result,cat_stock_growth_ratio_result = {},{},{},{}    
        stock_cat_pure_need,stock_cat_weighted_need = {},{}  
        stock_cat_pure_days,stock_cat_weighted_days = {},{}
        money_cat_ratio,money_cat_ratio_result = {},{}
        money_cat_ratio_perb,money_cat_ratio_perb_result = {},{}
        cat_money_weighted,cat_stock_weighted,cat_money_weighted_result,cat_stock_weighted_result = {},{},{},{}
        cat_stock_money_weighted,cat_stock_stock_weighted,cat_stock_money_weighted_result,cat_stock_stock_weighted_result = {},{},{},{}        
        cat_three_hold,cat_three_hold_result = {},{}
        rank_big_hold,rank_big_2_hold,cat_money_big,cat_stock_big,money_cat_ratio_big_perb,money_cat_ratio_big = {},{},{},{},{},{}
        money_cat_pure,money_cat_pure_result = {},{}
        
        for weight in weight_list :                               
            switch = 0 
            try:
                xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\資金觀察\\"+weight+"-資金.xls")
                xls_df_fund = xls_file.parse(xls_file.sheet_names[0], index_col=[0])   
            except:
                continue
            
            for k in range(len(xls_df_fund)):
                if int(xls_df_fund["日期"][k]) == int(xls_df_day["日期"][len(xls_df_day)-1]):
                    xls_df_fund = xls_df_fund[k:k+100]
                    xls_df_fund.index = range(len(xls_df_fund))
                    if len(xls_df_fund) >= 60 :
                        switch = 1
                    break
         
            xls_df_date = set(list(xls_df_fund["日期"]))
            xls_df_big_date = set(list(xls_df_day["日期"]))
            diff = xls_df_date.difference(xls_df_big_date)
            diff = list(diff)   
            if len(diff) != 0 : 
                diff.sort(reverse = False)
                for i in range(len(xls_df_fund)):
                    if xls_df_fund["日期"][i] in diff :
                        xls_df_fund.drop([i],axis = 0, inplace = True)
            xls_df_fund = xls_df_fund.sort_values(by = "日期",ascending = True)        
            xls_df_fund.index = range(len(xls_df_fund))

            if switch == 0:
                continue
            try:                 
                valid_list = []
                for k in range(1,22):
                    valid_list.append(float(xls_df_fund["股數"][len(xls_df_fund)-k]))
                for k in range(22,60):
                    if xls_df_fund["股數"][len(xls_df_fund)-k] >= xls_df_fund["股數"][len(xls_df_fund)-k-1] :
                        valid_list.append(float(xls_df_fund["股數"][len(xls_df_fund)-k]))
                    else:
                        break                 

                stock_cat_pure.setdefault(weight,valid_list)
                stock_cat_pure_days.setdefault(weight,len(valid_list))
            except:
                continue    

            try:
                valid_list = []
                for k in range(1,22):
                    valid_list.append(float(xls_df_day["股數值"][len(xls_df_day)-k]))
                for k in range(22,60):    
                    if xls_df_day["股數值"][len(xls_df_day)-k] >= xls_df_day["股數值"][len(xls_df_day)-k-1] :
                        valid_list.append(float(xls_df_day["股數值"][len(xls_df_day)-k]))
                    else:
                        break
                stock_cat_weighted.setdefault(weight,valid_list)
                stock_cat_weighted_days.setdefault(weight,len(valid_list))
            except:
                continue          

        for key in stock_cat_pure.keys():
            stock_cat_pure_result.setdefault(key,stock_cat_pure[key][0]-min(stock_cat_pure[key]))

        stock_cat_pure_5 = sorted(stock_cat_pure_result.items(),key=lambda x:x[1],reverse=True)[:5]

        for key in stock_cat_weighted.keys():
            stock_cat_weighted_result.setdefault(key,stock_cat_weighted[key][0]-min(stock_cat_weighted[key]))

        stock_cat_weighted_5 = sorted(stock_cat_weighted_result.items(),key=lambda x:x[1],reverse=True)[:5]
        
        valid_list_cat_pure = []
        for factor in stock_cat_pure_5 :
            valid_list_cat_pure.append(factor[0])

        valid_list_cat_weighted = []
        for factor in stock_cat_weighted_5 :
            valid_list_cat_weighted.append(factor[0])
            
        total_valid_cat = list(set(valid_list_cat_pure) & set(valid_list_cat_weighted))

        if len(total_valid_cat) != 0 :
            cat_big,cat_big_list = {},[] 
            for key in total_valid_cat:
                for j in range(930):
                    if table_industry.row_values(j)[2:3][0] == key:
                        big_cat = table_industry.row_values(j)[1:2][0]
                        cat_big.setdefault(key,big_cat) 
                        cat_big_list.append(big_cat)
                        break
            xls_df_basic["對應"][locate] = str(cat_big) 
            
        for weight in total_valid_cat :
            switch = 0 
            try:
                xls_file = pd.ExcelFile("D:\Stock Investment\股票數據\資金觀察"+weight+"-資金.xls")
                xls_df_fund = xls_file.parse(xls_file.sheet_names[0], index_col=[0])   
            except:
                continue
            
            for k in range(len(xls_df_fund)):
                if int(xls_df_fund["日期"][k]) == int(xls_df_day["日期"][len(xls_df_day)-1]):
                    xls_df_fund = xls_df_fund[k:k+100]
                    xls_df_fund.index = range(len(xls_df_fund))
                    if len(xls_df_fund) >= 60 :
                        switch = 1
                    break
                
            xls_df_date = set(list(xls_df_fund["日期"]))
            xls_df_big_date = set(list(xls_df_day["日期"]))
            diff = xls_df_date.difference(xls_df_big_date)
            diff = list(diff) 
            if len(diff) != 0 : 
                diff.sort(reverse = False)
                for i in range(len(xls_df_fund)):
                    if xls_df_fund["日期"][i] in diff :
                        xls_df_fund.drop([i],axis = 0, inplace = True)
            xls_df_fund = xls_df_fund.sort_values(by = "日期",ascending = True)        
            xls_df_fund.index = range(len(xls_df_fund))
            
            if switch == 0:
                continue
            
            if weight not in stock_cat_pure.keys() :
                days_cat = stock_cat_weighted_days[weight]
                
                days_need = need_days(xls_df_fund, days = days_cat,list_need = ["股數"])            
                try:                   
                    valid_list = []
                    for k in range(1,days_need+1):
                        valid_list.append(float(xls_df_fund["股數"][len(xls_df_fund)-k]))  
                        
                    stock_cat_pure.setdefault(weight,valid_list)                                                                   
                except:
                    if len(valid_list) != 0 :
                        stock_cat_pure.setdefault(weight,valid_list) 
                        pass
                    else:
                        continue    

            if weight in stock_cat_weighted.keys() and len(stock_cat_pure[weight]) > len(stock_cat_weighted[weight]):
                days_cat = len(stock_cat_pure[weight]) - len(stock_cat_weighted[weight])
                try:
                    n = 0
                    for k in range(1+len(stock_cat_weighted[weight]),30):
                        if xls_df_day["股數值"][len(xls_df_day)-k] == 100 :
                            stock_cat_weighted[weight].append(float(xls_df_day["股數值"][len(xls_df_day)-k]))
                        elif xls_df_day["股數值"][len(xls_df_day)-k] > xls_df_day["股數值"][len(xls_df_day)-k-1] :
                            stock_cat_weighted[weight].append(float(xls_df_day["股數值"][len(xls_df_day)-k]))
                        else:
                            stock_cat_weighted[weight].append(float(xls_df_day["股數值"][len(xls_df_day)-k]))
                            n = k
                            break

                    for k in range(1+n,60):
                        if len(stock_cat_weighted[weight]) < len(stock_cat_pure[weight]):
                            stock_cat_weighted[weight].append(float(xls_df_day["股數值"][len(xls_df_day)-k]))  
                        else:
                            break

                except:
                    if len(valid_list) != 0 :
                        stock_cat_weighted.setdefault(weight,valid_list) 
                        pass
                    else:
                        continue   
           
                
            elif weight not in stock_cat_weighted.keys() :     
                days_cat = stock_cat_pure_days[weight]
                
                days_need = need_days(xls_df_day, days = days_cat,list_need = ["股數值"])   
                try:
                    valid_list= []
                    for k in range(1,days_need+1):
                        valid_list.append(float(xls_df_day["股數值"][len(xls_df_day)-k]))
                      
                    stock_cat_weighted.setdefault(weight,valid_list)
                except:
                    if len(valid_list) != 0 :
                        stock_cat_weighted.setdefault(weight,valid_list) 
                        pass
                    else:
                        continue       

            if weight in stock_cat_pure_days.keys() and weight in stock_cat_weighted_days.keys() :
                days_cat = stock_cat_pure_days[weight]             
            elif weight in stock_cat_pure_days.keys() and weight not in stock_cat_weighted_days.keys() :
                days_cat = stock_cat_pure_days[weight]
            elif weight not in stock_cat_pure_days.keys() and weight in stock_cat_weighted_days.keys() :
                days_cat = stock_cat_weighted_days[weight]
 
            days_need = need_days(xls_df_day, days = days_cat,list_need = ["weight股數值"])                                    
            valid_list = []
            for k in range(1,days_need+1):
                valid_list.append(float(xls_df_day["weight股數值"][len(xls_df_day)-k]))                      
            stock_cat_stock.setdefault(weight,valid_list)                                        
 
    
            days_need = need_days(xls_df_day, days = days_cat,list_need = ["weight股數比"])                                
            valid_list = []
            for k in range(1,days_need+1):
                valid_list.append(float(xls_df_day["weight股數比"][len(xls_df_day)-k]))            
            stock_stock_pure.setdefault(weight,valid_list)                   


            days_need = need_days(xls_df_fund, days = days_cat,list_need = ["金額"])                                    
            valid_list= []
            for k in range(1,days_need+1):
                valid_list.append(float(xls_df_fund["金額"][len(xls_df_fund)-k]))           
            money_cat_pure.setdefault(weight,valid_list)          

            days_need = need_days(xls_df_fund, days = days_cat,list_need = ["股數增減率"])                                    
            valid_list= []
            for k in range(1,days_need+1):
                valid_list.append(float(xls_df_fund["股數增減率"][len(xls_df_fund)-k]))           
            cat_growth_perb.setdefault(weight,valid_list)     

            days_need = need_days(xls_df_day, days = days_cat,list_need = ["weight股數增減率","weight+股數增減率值"])                                    
            valid_list_s,valid_list_m = [],[]
            for k in range(1,days_need+1):
                valid_list_s.append(float(xls_df_day["weight股數增減率"][len(xls_df_day)-k]))       
                valid_list_m.append(float(xls_df_day["weight+股數增減率值"][len(xls_df_day)-k]))    
            cat_stock_growth_ratio.setdefault(weight,valid_list_s) 
            cat_stock_growth.setdefault(weight,valid_list_m) 


            days_need = need_days_pair(xls_df_day,days_cat,list_need = ["weight股數","weight金額"])                   
            valid_list_s,valid_list_m = [],[]
            for k in range(1,days_need+1):
                valid_list_s.append(float(xls_df_day["weight股數"][len(xls_df_day)-k]))
                valid_list_m.append(float(xls_df_day["weight金額"][len(xls_df_day)-k])) 
            cat_stock_pure_.setdefault(weight,valid_list_s)
            cat_money_pure_.setdefault(weight,valid_list_m)                  
           
            
            days_need = need_days_pair(xls_df_fund, days = days_cat,list_need = ["金額值","股數值"])                    
            valid_list_s,valid_list_m = [],[]
            for k in range(1,days_need+1):
                valid_list_s.append(float(xls_df_fund["股數值"][len(xls_df_fund)-k]))
                valid_list_m.append(float(xls_df_fund["金額值"][len(xls_df_fund)-k]))   
            cat_stock_.setdefault(weight,valid_list_s)             
            cat_money_.setdefault(weight,valid_list_m)                 


            days_need = need_days_pair(xls_df_fund, days = days_cat,list_need = ["金額比例","金額比例值"])    
            valid_list_s,valid_list_m = [],[]
            for k in range(1,days_need+1):
                valid_list_s.append(float(xls_df_fund["金額比例"][len(xls_df_fund)-k]))
                valid_list_m.append(float(xls_df_fund["金額比例值"][len(xls_df_fund)-k]))
            money_cat_ratio.setdefault(weight,valid_list_s)
            money_cat_ratio_perb.setdefault(weight,valid_list_m)                  
           

            days_need = need_days_pair(xls_df_day, days = days_cat,list_need = ["weight金額值_1","weight股數值_1"])                      
            valid_list_s,valid_list_m = [],[]
            for k in range(1,days_need+1):
                valid_list_s.append(float(xls_df_day["weight股數值_1"][len(xls_df_day)-k]))
                valid_list_m.append(float(xls_df_day["weight金額值_1"][len(xls_df_day)-k]))         
            cat_stock_weighted.setdefault(weight,valid_list_s) 
            cat_money_weighted.setdefault(weight,valid_list_m)                 
           
            
            days_need = need_days_pair(xls_df_day, days = days_cat,list_need = ["金額_1","股數_1"])            
            valid_list_s,valid_list_m = [],[]
            for k in range(1,days_need+1):  
                valid_list_s.append(float(xls_df_day["金額_1"][len(xls_df_day)-k]))
                valid_list_m.append(float(xls_df_day["股數_1"][len(xls_df_day)-k]))               
            cat_stock_stock_weighted.setdefault(weight,valid_list_s)  
            cat_stock_money_weighted.setdefault(weight,valid_list_m)   
              
            for j in range(930):
                if table_industry.row_values(j)[2:3][0] == weight:
                    big_cat = table_industry.row_values(j)[1:2][0]
                    break
                    
            xls_file = pd.ExcelFile("D:\\Stock Investment\\股票數據\\資金觀察\\"+big_cat+"-巨資金.xls")
            xls_df_fund = xls_file.parse(xls_file.sheet_names[0], index_col=[0])   
            for k in range(len(xls_df_fund)):
                if int(xls_df_fund["日期"][k]) == int(xls_df_day["日期"][len(xls_df_day)-1]):
                    xls_df_fund = xls_df_fund[k:k+100]
                    xls_df_fund.index = range(len(xls_df_fund))
                    break 
                
            xls_df_date = set(list(xls_df_fund["日期"]))
            xls_df_big_date = set(list(xls_df_day["日期"]))
            diff = xls_df_date.difference(xls_df_big_date)
            diff = list(diff)   
            if len(diff) != 0 : 
                diff.sort(reverse = False)
                for i in range(len(xls_df_fund)):
                    if xls_df_fund["日期"][i] in diff :
                        xls_df_fund.drop([i],axis = 0, inplace = True)                
            xls_df_fund = xls_df_fund.sort_values(by = "日期",ascending = True)   
            xls_df_fund.index = range(len(xls_df_fund)) 
            
            
            days_need = need_days_pair(xls_df_fund, days = days_cat,list_need = ["金額值","股數值"])                    
            valid_list_s,valid_list_m = [],[]
            for k in range(1,days_need+1):
                valid_list_s.append(float(xls_df_fund["股數值"][len(xls_df_fund)-k])) 
                valid_list_m.append(float(xls_df_fund["金額值"][len(xls_df_fund)-k]))   
            cat_stock_big.setdefault(weight,valid_list_s)                  
            cat_money_big.setdefault(weight,valid_list_m)                 

            days_need = need_days(xls_df_fund, days = days_cat,list_need = ["股數增減率"])                  
            valid_list_s = []
            for k in range(1,days_need+1):
                valid_list_s.append(float(xls_df_fund["股數增減率"][len(xls_df_fund)-k]))       
            cat_growth_big_perb.setdefault(weight,valid_list_s)  
        
            days_need = need_days_pair(xls_df_fund, days = days_cat,list_need = ["金額比例","金額比例值"])                                  
            valid_list_s,valid_list_m = [],[]
            for k in range(1,days_need+1):
                valid_list_s.append(float(xls_df_fund["金額比例"][len(xls_df_fund)-k]))
                valid_list_m.append(float(xls_df_fund["金額比例值"][len(xls_df_fund)-k]))   
            money_cat_ratio_big.setdefault(weight,valid_list_s)                            
            money_cat_ratio_big_perb.setdefault(weight,valid_list_m)                 


            rank_big_hold.setdefault(weight,xls_df_fund["股數"][len(xls_df_fund)-1]) 
            rank_big_2_hold.setdefault(weight,xls_df_fund["股數"][len(xls_df_fund)-2]) 
                        
            xls_df_big.index = range(len(xls_df_big))   
            xls_df_date = set(list(xls_df_big["日期"]))
            xls_df_big_date = set(list(xls_df_day["日期"]))
            diff = xls_df_date.difference(xls_df_big_date)
            diff = list(diff)   
            if len(diff) != 0 :
                diff.sort(reverse = False)
                for i in range(len(xls_df_big)):
                    if xls_df_big["日期"][i] in diff :
                        xls_df_big.drop([i],axis = 0, inplace = True)
            xls_df_big.index = range(len(xls_df_big))
            
        for weight in total_valid_cat:
            stock_cat_stock_result.setdefault(weight,stock_cat_stock[weight])            
            stock_stock_pure_result.setdefault(weight,stock_stock_pure[weight])
            stock_cat_pure_need.setdefault(weight,stock_cat_pure[weight])            
            stock_cat_weighted_need.setdefault(weight,stock_cat_weighted[weight])
            money_cat_pure_result.setdefault(weight,money_cat_pure[weight])
            cat_stock_pure_result.setdefault(weight,cat_stock_pure_[weight])
            cat_money_pure_result.setdefault(weight,cat_money_pure_[weight])
            cat_money_result.setdefault(weight,cat_money_[weight])
            cat_stock_result.setdefault(weight,cat_stock_[weight])
            money_cat_ratio_result.setdefault(weight,money_cat_ratio[weight])
            money_cat_ratio_perb_result.setdefault(weight,money_cat_ratio_perb[weight])
            cat_money_weighted_result.setdefault(weight,cat_money_weighted[weight])
            cat_stock_weighted_result.setdefault(weight,cat_stock_weighted[weight])    
            cat_stock_money_weighted_result.setdefault(weight,cat_stock_money_weighted[weight]) 
            cat_stock_stock_weighted_result.setdefault(weight,cat_stock_stock_weighted[weight])   
            cat_three_hold_result.setdefault(weight,cat_three_hold[weight]) 
            cat_stock_growth_result.setdefault(weight,cat_stock_growth[weight]) 
            cat_stock_growth_ratio_result.setdefault(weight,cat_stock_growth_ratio[weight]) 
            cat_growth_perb_result.setdefault(weight,cat_growth_perb[weight])
            cat_growth_big_perb_result.setdefault(weight,cat_growth_big_perb[weight])    
            cat_uni_perb_result.setdefault(weight,cat_uni_perb[weight])    
            cat_price_perb_result.setdefault(weight,cat_price_perb[weight])    
            cat_uni_big_result.setdefault(weight,cat_uni_big[weight])    
            cat_price_big_result.setdefault(weight,cat_price_big[weight])    
            

        xls_df_basic["金額"][locate] = str(money_cat_pure_result)
        xls_df_basic["金額比例"][locate] = str(money_cat_ratio_result)
        xls_df_basic["金額比例值"][locate] = str(money_cat_ratio_perb_result)
        xls_df_basic["股數"][locate] = str(cat_stock_pure_result)
        xls_df_basic["金額_1"][locate] = str(cat_money_pure_result)
        xls_df_basic["股數_1"][locate] = str(cat_stock_result)
        xls_df_basic["金額比"][locate] = str(cat_money_result)
        xls_df_basic["股數比"][locate] = str(cat_stock_weighted_result)
        xls_df_basic["金額比值"][locate] = str(cat_money_weighted_result)
        xls_df_basic["股數比值"][locate] = str(cat_stock_stock_weighted_result)
        xls_df_basic["金額比值_1"][locate] = str(cat_stock_money_weighted_result)
        xls_df_basic["股數增長率"][locate] = str(cat_stock_growth_result)
        xls_df_basic["股數增長率比"][locate] = str(cat_stock_growth_ratio_result)
        xls_df_basic["法人持有"][locate] = str(cat_three_hold_result)
        xls_df_basic["持有-今日"][locate] = str(rank_big_hold)
        xls_df_basic["持有-昨日"][locate] = str(rank_big_2_hold)  
        xls_df_basic["金額_2"][locate] = str(cat_stock_big)  
        xls_df_basic["金額_3"][locate] = str(cat_money_big) 
        xls_df_basic["金額比例"][locate] = str(money_cat_ratio_big)  
        xls_df_basic["金額比例值"][locate] = str(money_cat_ratio_big_perb) 
        xls_df_basic["股值"][locate] = str(total_valid_cat)
        xls_df_basic["增長率"][locate] = str(cat_growth_perb_result)
        xls_df_basic["增長率_1"][locate] = str(cat_growth_big_perb_result)
