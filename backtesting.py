# -*- coding: utf-8 -*-
"""
Created on Sun Jan  9 14:27:32 2022

@author: yao79
"""

import numpy as np
import pandas as pd

xls_file_big_d = pd.ExcelFile("D:\\Stock Investment\\股票數據\\市盤.xlsx")
xls_file_big_d = xls_file_big_d.parse(xls_file_big_d.sheet_names[0], index_col=[0])                
date_all = np.array(xls_file_big_d["日期"]).tolist()
date_all.sort()
target_date = 20170101
day_dic = {}
trade_list = []
                                    
category,times,stock_num = 0,0,0 
for date in date_all:
    if date < target_date or date >= 20180101:
        continue
###讀取候選檔案
    try:
        xls_df_basic = pd.ExcelFile(r"C:\\Users\\User\\Desktop\\每日計算報告\\市候選"+str(date)+".xls")
        xls_df_basic = xls_df_basic.parse(xls_df_basic.sheet_names[0], index_col=[0]) 
    except:
        print(str(date))
        continue
###讀取當天候選個股
    candidate_list = []   
    for i in range(len(xls_df_basic)):
        candidate_list.append(xls_df_basic["股票名稱"][i])
###進入回測環潔        
    for stock in candidate_list:
###讀取該股日報表                        
        xls_df_day = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report2015\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)   
        buy_price,switch = 0,0
###時間大於抓到日開始進行後續股價驗證        
        for k in range(len(xls_df_day)):
            try:###購買後 buy_price != 0，便不在執行此流程
                if int(xls_df_day["日期"][k]) > date and buy_price == 0:    
                    ###抓到日隔天是否有碰到股價購買條件
                    if (xls_df_day["日最高價"][k] - xls_df_day["日收盤價"][k-1])/xls_df_day["日收盤價"][k-1] > 0.01 and buy_price == 0 :
                        if xls_df_day["日最高價"][k] >= max(xls_df_day["日開盤價"][k-5:k]) and xls_df_day["日最高價"][k] >= max(xls_df_day["日收盤價"][k-5:k]) and (xls_df_day["日最高價"][k] - max(xls_df_day["日最高價"][k-5:k])) / xls_df_day["日最高價"][k] > 0.01 :              
                            if xls_df_day["日開盤價"][k] > max(xls_df_day["日最高價"][k-5:k]) * 1.01 :
                                buy_price =  xls_df_day["日開盤價"][k] * 1000
                                risk_price = (xls_df_day["日最高價"][k-1] + xls_df_day["日最低價"][k-1]) / 2
                            elif xls_df_day["日開盤價"][k] <= max(xls_df_day["日最高價"][k-5:k]) * 1.01 <= xls_df_day["日最高價"][k]:
                                buy_price =  max(xls_df_day["日最高價"][k-5:k]) * 1.01 * 1000
                                risk_price = (xls_df_day["日最高價"][k-1] + xls_df_day["日最低價"][k-1]) / 2
                            
                            buy_day = k
                            if buy_price * 0.001425 >= 20 :
                                fee_buy = buy_price * 0.001425
                            else:
                                fee_buy = 20
                        else:
                            break
                    ###若隔天沒有買入，則直接放棄個股，換下一檔候選股驗證                    
                    elif buy_price == 0 :
                        break

  ###購買後 buy_price != 0，只會執行此流程；
  ###驗證購買個股是否有碰到停損 / 停利點，碰觸之後看隔天表現，若仍維持停損 / 停利條件，則賣出
                if buy_price!= 0 and xls_df_day["日收盤價"][k] <= risk_price and switch == 0:
                    switch = 1
                    continue                
                elif buy_price!= 0 and xls_df_day["日收盤價"][k] <= xls_df_day["10日均線"][k] and switch == 0:
                    switch = 1
                    continue
                elif  buy_price!= 0 and xls_df_day["日收盤價"][k] <= 0.5*(xls_df_day["日收盤價"][buy_day]+xls_df_day["日開盤價"][buy_day]) and k - buy_day >= 5 and switch == 0:
                    switch = 1
                    continue
                if switch == 1 :
                    if xls_df_day["日收盤價"][k] <= risk_price or xls_df_day["日收盤價"][k] <= xls_df_day["10日均線"][k]:        
                        sell_price = xls_df_day["日收盤價"][k] * 1000
                        if sell_price * 0.001425 >= 20 :
                            fee_sell = sell_price * 0.001425
                        else:
                            fee_sell = 20
                            
                        tax = sell_price * 0.003
                        profit = sell_price - buy_price - fee_buy - tax - fee_sell
                        print(stock,str(profit),str(k-buy_day)+"天")
                        if stock in day_dic.keys():
                            day_dic[stock] = int(xls_df_day["日期"][k])
                        else:
                           day_dic.setdefault(stock,int(xls_df_day["日期"][k]))
                         
                        trade_list.append([stock,round(profit,2),str(k-buy_day)+"天",str(buy_price),str(fee_buy),str(sell_price),str(fee_sell),str(tax)])
                
            except:
                continue
               
trade_list = pd.DataFrame(trade_list,columns = ["股票名稱","獲利","持有天數","買價","買手續費","賣價","賣手續費","交易稅"])
trade_list.to_excel("C:\\Users\\User\\Desktop\\交易清單-市2017.xlsx")                   
