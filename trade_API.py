# -*- coding: utf-8 -*-
"""
Created on Sun Nov  5 21:21:07 2023

@author: User
"""

from fugle_realtime import HttpClient
from configparser import ConfigParser
from fugle_trade.sdk import SDK
from fugle_trade.order import OrderObject
from fugle_trade.constant import (APCode, Trade, PriceFlag, BSFlag, Action)
import pandas as pd
import requests
import time
import datetime
import math

switch = 0
while switch == 0:
    if  int(str(datetime.datetime.now())[11:13]+str(datetime.datetime.now())[14:16]) >= 901 :    
        switch = 1
    else:
        time.sleep(60)
        

# 載入候選名單
xls_file_candidate = pd.ExcelFile("D:\\Stock Investment\\交易檔案\\候選表.xlsx")
xls_df_candidate = xls_file_candidate.parse(xls_file_candidate.sheet_names[0], index_col=[0])

xls_file_stock = pd.ExcelFile("D:\\Stock Investment\\交易檔案\\庫存表.xlsx")
xls_df_stock = xls_file_stock.parse(xls_file_stock.sheet_names[0], index_col=[0])

token = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
headers = { "Authorization": "Bearer " + token }
api_client = HttpClient(api_token='xxxxxxxxxxxxxxxxxxxxxxxxxxxx')

cost,count,all_buy_list,all_sell_list,limit = 0,0,[],[],370000

today_std = int((str(datetime.datetime.now())[:4]+str(datetime.datetime.now())[5:7]+str(datetime.datetime.now())[8:10]))

while 900 <= int(str(datetime.datetime.now())[11:13]+str(datetime.datetime.now())[14:16]) <= 1325:
    buy_list,sell_list,switch_b,switch_s,switch_t = [],[],0,0,0
    for k in range(len(xls_df_candidate)):
        if math.isnan(xls_df_candidate["股票代號"][k]):
            continue
        try:
            price = api_client.intraday.quote(symbolId=xls_df_candidate["股票代號"][k])["data"]["quote"]["trade"]["price"]
            time_at = api_client.intraday.quote(symbolId=xls_df_candidate["股票代號"][k])['data']["info"]['lastUpdatedAt']
        except:
            continue
        print(str(price)+","+str(time_at), flush = True)
        if cost <= limit and price >= xls_df_candidate["購買價"][k] and xls_df_candidate["預約單完成"][k] == 0 and xls_df_candidate["尾盤購買"][k] == 0:
            order = OrderObject(
                ap_code = APCode.Common,
                trade = Trade.Cash,
                buy_sell = Action.Buy,    
                price_flag = PriceFlag.Market,
                price = None,
                stock_no = str(int(xls_df_candidate["股票代號"][k])),
                quantity = int(xls_df_candidate["買進張數"][k]),
                bs_flag = BSFlag.ROD)            
            
            sdk.place_order(order)            
            buy_list.append(int(xls_df_candidate["股票代號"][k]))
            
            cost = cost + price*1000 

        elif cost <= limit and xls_df_candidate["尾盤購買"][k] == 1 and xls_df_candidate["預約單完成"][k] == 0 and api_client.intraday.quote(symbolId=xls_df_candidate["股票代號"][k])["data"]["quote"]["changePercent"] >= 9 :
            order = OrderObject(
                ap_code = APCode.Common,
                trade = Trade.Cash,
                buy_sell = Action.Buy,    
                price_flag = PriceFlag.Market,
                price = None,
                stock_no = str(int(xls_df_candidate["股票代號"][k])),
                quantity = int(xls_df_candidate["買進張數"][k]),
                bs_flag = BSFlag.ROD)            
            
            sdk.place_order(order)            
            buy_list.append(int(xls_df_candidate["股票代號"][k]))
            
            cost = cost + price*1000             
                        
        elif 1323 <= int(str(datetime.datetime.now())[11:13]+str(datetime.datetime.now())[14:16]) <= 1325:
            switch_stock = 0
            if cost <= limit and price >= xls_df_candidate["購買價"][k] and xls_df_candidate["預約單完成"][k] == 0 and xls_df_candidate["尾盤購買"][k] == 1 :
                priceHigh = api_client.intraday.quote(symbolId=xls_df_candidate["股票代號"][k])["data"]["quote"]["priceHigh"]["price"]
                priceLow = api_client.intraday.quote(symbolId=xls_df_candidate["股票代號"][k])["data"]["quote"]["priceLow"]["price"]
                priceOpen = api_client.intraday.quote(symbolId=xls_df_candidate["股票代號"][k])["data"]["quote"]["priceOpen"]["price"]
                amplitude = api_client.intraday.quote(symbolId=xls_df_candidate["股票代號"][k])["data"]["quote"]["amplitude"]
                if (price - priceOpen) > 0 and (priceHigh - price) / (priceHigh - priceLow) <= 0.33 : 
                    switch_stock = 1
                elif (price - priceOpen) < 0 and (price - priceLow) / (priceHigh - priceLow) >= 0.67 :                     
                    switch_stock = 1
                if (price - priceOpen) > 0 and (price - priceOpen)/(priceHigh - priceLow) <= 0.4 and (priceHigh - price) / (priceHigh - priceLow) >= 0.33 and (priceOpen - priceLow) / (priceHigh - priceLow) >= 0.2:                     
                    switch_stock = 0
                elif (price - priceOpen) < 0 and (priceOpen - price)/(priceHigh - priceLow) <= 0.4 and (priceHigh - priceOpen) / (priceHigh - priceLow) >= 0.2 and (price - priceLow) / (priceHigh - priceLow) >= 0.33:                     
                    switch_stock = 0                  
                    
                if switch_stock == 1 :
                    order = OrderObject(
                        ap_code = APCode.Common,
                        trade = Trade.Cash,
                        buy_sell = Action.Buy,    
                        price_flag = PriceFlag.Market,
                        price = None,
                        stock_no = str(int(xls_df_candidate["股票代號"][k])),
                        quantity = int(xls_df_candidate["買進張數"][k]),
                        bs_flag = BSFlag.ROD)   
                    
                    sdk.place_order(order)            
                    buy_list.append(int(xls_df_candidate["股票代號"][k]))
            
                    cost = cost + price*1000 
                
    if 1323 <= int(str(datetime.datetime.now())[11:13]+str(datetime.datetime.now())[14:16]) <= 1325:    
        for k in range(len(xls_df_stock)):
            if xls_df_stock["賣出完成"][k] == 1 or xls_df_stock["10日均線"][k] == "":
                continue
            switch_stock = 0
            price = api_client.intraday.quote(symbolId=xls_df_stock["股票代號"][k])["data"]["quote"]["trade"]["price"]    
            if price <= xls_df_stock["10日均線"][k] and xls_df_stock["跌破10日均線次數"][k] >= 1 and xls_df_stock["預約單完成"][k] == 0:
                switch_stock = 1
            elif price <= xls_df_stock["停損價"][k] and xls_df_stock["跌破停損價次數"][k] >= 1 and xls_df_stock["預約單完成"][k] == 0:
                switch_stock = 1         
            elif price <= xls_df_stock["當日成本"][k] and xls_df_stock["持有天數"][k] >= 6 and xls_df_stock["跌破當日成本次數"][k] >= 1 and xls_df_stock["預約單完成"][k] == 0:
                switch_stock = 1 
                
            if switch_stock == 1 :
                order = OrderObject(
                    ap_code = APCode.Common,
                    trade = Trade.Cash,
                    buy_sell = Action.Sell,    
                    price_flag = PriceFlag.Market,
                    price = None,
                    stock_no = str(int(xls_df_stock["股票代號"][k])),
                    quantity = int(xls_df_stock["持有張數"][k]),
                    bs_flag = BSFlag.ROD)          
                
                sdk.place_order(order) 
                sell_list.append(int(xls_df_stock["股票代號"][k]))
                
                                
    order_results = sdk.get_order_results()
    for result in order_results:
        for stock in buy_list:  
            if result["stock_no"] == str(stock) :
                for k in range(len(xls_df_candidate)):
                    if xls_df_candidate["股票代號"][k] == stock and xls_df_candidate["預約單完成"][k] == 0:
                        xls_df_candidate["預約單完成"][k] = 1
                        switch_b = 1
                        break
                        
            if result["stock_no"] == str(stock) and result["mat_qty"] != 0 and result["buy_sell"] == "B" :
                for k in range(len(xls_df_candidate)):
                    if xls_df_candidate["股票代號"][k] == stock and xls_df_candidate["買進完成"][k] == 0 :
                        xls_df_candidate["買進完成"][k] = 1
                        stock = xls_df_candidate["股票名稱"][k] 
                        if stock not in all_buy_list:
                            data = { 'message': stock + "已經完成買進，成本為" + str(result["avg_price"])}
                            requests.post("https://notify-api.line.me/api/notify",headers = headers, data = data)
                            all_buy_list.append(int(xls_df_candidate["股票代號"][k]))
                            switch_t = 1

                            info = [today_std,result["mat_qty"],"",result["stock_no"],stock,result["avg_price"],"","","","","","","",0,0,"",""]
                            info = pd.DataFrame([info],columns = xls_df_stock.columns)
                            below = xls_df_stock.loc[:]
                            xls_df_stock = info.append(below,ignore_index=True)   
                            break
                            
        for stock in sell_list:  
            if result["stock_no"] == str(stock) :
                for k in range(len(xls_df_stock)):
                    if xls_df_stock["股票代號"][k] == stock and xls_df_stock["預約單完成"][k] == 0:
                        xls_df_stock["預約單完成"][k] = 1
                        switch_s = 1
                        break
                        
            if result["stock_no"] == str(stock) and result["mat_qty"] != 0 and result["buy_sell"] == "S":
                for k in range(len(xls_df_stock)):
                    if xls_df_stock["股票代號"][k] == stock and xls_df_stock["賣出完成"][k] == 0:
                        xls_df_stock["賣出完成"][k] = 1
                        xls_df_stock["賣出日期"][k] = today_std      
                        xls_df_stock["賣出價"][k] = result["avg_price"]
                        stock = xls_df_stock["股票名稱"][k] 
                        
                        if result["avg_price"]* 1000 * 0.001425 >= 20 :
                            fee_sell = result["avg_price"] * 0.001425 * 1000
                        else:
                            fee_sell = 20

                        if xls_df_stock["購買價"][k] * 1000 * 0.001425 >= 20 :
                            fee_buy = xls_df_stock["購買價"][k] * 0.001425 * 1000
                        else:
                            fee_buy = 20

                        tax = result["avg_price"] * 0.003 * 1000
                        xls_df_stock["獲利"][k] = round(result["avg_price"] * 1000 - xls_df_stock["購買價"][k] * 1000 - fee_buy - tax - fee_sell,2)                                                                                 

                        if stock not in all_sell_list:
                            data = { 'message': stock + "已經完成賣出，賣價為" + str(result["avg_price"])+"，獲利為"+str(xls_df_stock["獲利"][k])}
                            requests.post("https://notify-api.line.me/api/notify",headers = headers, data = data)
                            all_sell_list.append(int(xls_df_stock["股票代號"][k]))
                            switch_s = 1
                            break
                            
                            
    if switch_b == 1 and switch_t == 0:
        xls_df_candidate.to_excel("D:\\Stock Investment\\交易檔案\\候選表.xlsx")                    
       
    if switch_t == 1 :
        xls_df_candidate.to_excel("D:\\Stock Investment\\交易檔案\\候選表.xlsx")                    
        xls_df_stock.to_excel("D:\\Stock Investment\\交易檔案\\庫存表.xlsx") 
        
    if switch_s == 1 and switch_t == 0 :                 
        xls_df_stock.to_excel("D:\\Stock Investment\\交易檔案\\庫存表.xlsx")                                          
                                         
    time.sleep(30)     
