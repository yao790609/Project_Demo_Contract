# -*- coding: utf-8 -*-

"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import matplotlib.pyplot as plt
import xlrd
import pic_module
import mpl_finance as mpf

def qualified_day(start_day,end_day,xls_df_day_1):
    global total_all_qualified
    block,total,total_all_qualified = [],[],[]
    for i in range(start_day,end_day):
        if xls_df_day_1["日收盤價"][i] > xls_df_day_1["日上軌"][i]:
            block.append(str(int(xls_df_day_1["日期"][i])))
            if xls_df_day_1["動能指標"][i] <= 3 and xls_df_day_1["動能指標_1"][i] <= 3 and xls_df_day_1["動能指標_2"][i] <= 3 :
                if xls_df_day_1["動能指標_3"][i] <= 3 and xls_df_day_1["動能指標_4"][i] <= 3 and xls_df_day_1["動能指標_5"][i] <= 3 :            
                    total_all_qualified.append(xls_df_day_1["日期"][i])

        else:
            if len(block) == 0 :
                continue
            else:
                total.append([block[0],block[-1]])
                block = []

    if len(block) != 0:
        total.append([block[0],block[-1]])
        block = []
    for j in total :
        plt.axvspan(j[0],j[-1],facecolor = "grey")

    return total_all_qualified   

#上市    
file = xlrd.open_workbook(r"D:\Stock Investment\股票分類-下載用(去除重複)-上市.xlsx")
table = file.sheets()[0]
sheet = file.sheets()[1]

for i in range(33):
    stock_list = table.row_values(i)[2:]
    number_list = sheet.row_values(i)[2:]
    for stock in stock_list:
        data_list = []
        if stock != '' : 
                    
            xls_df_day = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report_KD\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)
            xls_df_day_1 = pd.read_csv("C:\\Users\\User\\Desktop\\Daily Report\\"+stock+"-day.csv",engine='python', encoding = "big5",index_col=0)

            for k in range(len(xls_df_day)):
                if int(xls_df_day["日期"][k]) < int(xls_df_day_1["日期"][0]):
                    xls_df_day.drop([k],axis = 0, inplace = True) 
            xls_df_day.index = range(len(xls_df_day))
            
            for k in range(len(xls_df_day)):
                if int(xls_df_day["日期"][k]) > int(xls_df_day_1["日期"][len(xls_df_day_1) - 1]):
                    xls_df_day.drop([k],axis = 0, inplace = True)    
            xls_df_day.index = range(len(xls_df_day))      
                  
            start_day = len(xls_df_day) - 250
            end_day = len(xls_df_day) - 0

            if start_day < 0:
                start_day = 0
            if end_day > len(xls_df_day):
                end_day = len(xls_df_day)


            ###個股日

            print(stock+"日")
            plt.figure(figsize=(30,60))

            
            x = []
            if len(x) == 0:
                for i in range(start_day,end_day):
                    x.append(str(int(xls_df_day["日期"][i])))   

            ax3 = plt.subplot2grid((25, 2), (0, 0), colspan=4,rowspan=2)
            plt.grid(True) 
      
    
            y4 = xls_df_day_1["日上軌"][start_day:end_day]      
            y5 = xls_df_day_1["日中軸"][start_day:end_day]      
            y6 = xls_df_day_1["日下軌"][start_day:end_day]  
            y7 = xls_df_day_1["60日均線"][start_day:end_day]
            y8 = xls_df_day_1["5日均線"][start_day:end_day]  
            y9 = xls_df_day_1["10日均線"][start_day:end_day]

            plt.ylim(min(y6), max(y4))                                  
            ax3.plot(x,y4,color='black')
            ax3.plot(x,y5,color='#FF9933')
            ax3.plot(x,y6,color='black')
            ax3.plot(x,y7,color='#00DD00')
            ax3.plot(x,y8,color='#668800')
            ax3.plot(x,y9,color='#FF00FF')     
            ax3.yaxis.set_label_position("left")
            qualified_day(start_day,end_day,xls_df_day_1)
            plt.legend()            
        
            
            ax4 = ax3.twinx() 
            mpf.candlestick2_ochl(ax4, xls_df_day["日開盤價"][start_day:end_day] ,xls_df_day["日收盤價"][start_day:end_day] , xls_df_day["日最高價"][start_day:end_day] ,xls_df_day["日最低價"][start_day:end_day] , width=0.6, colorup='r', colordown='g', alpha=0.75)
            y16 = xls_df_day_1["日上軌"][start_day:end_day]
            y17 = xls_df_day_1["日下軌"][start_day:end_day]
    
            plt.ylim(min(y17), max(y16))   
            ax4.plot(x,y17,color='black')
            ax4.plot(x,y16,color='black')
            plt.legend()
        
            ax5 = plt.subplot2grid((25, 2), (2, 0), colspan=4,rowspan=2)
            plt.grid(True) 
            
            y1 = xls_df_day_1["日收盤價"][start_day:end_day]      
            y4 = xls_df_day_1["日上軌"][start_day:end_day]      
            y5 = xls_df_day_1["日中軸"][start_day:end_day]      
            y6 = xls_df_day_1["日下軌"][start_day:end_day]  
            y7 = xls_df_day_1["60日均線"][start_day:end_day]
            y8 = xls_df_day_1["5日均線"][start_day:end_day]  
            y9 = xls_df_day_1["10日均線"][start_day:end_day]
                
            ax5.plot(x,y1,color='blue')                    
            ax5.plot(x,y4,color='black')
            ax5.plot(x,y5,color='#FF9933')
            ax5.plot(x,y6,color='black')
            ax5.plot(x,y7,color='#00DD00')
            ax5.plot(x,y8,color='#668800')
            ax5.plot(x,y9,color='#FF00FF')     
            ax5.yaxis.set_label_position("left")
            qualified_day(start_day,end_day,xls_df_day_1)
            plt.legend()        

            ax6 = ax5.twinx() 
            y1 = xls_df_day_1["動能指標"][start_day:end_day]   
            y2 = xls_df_day_1["動能指標_1"][start_day:end_day] 
            ax6.plot(x,y1,color='#6A0DAD',lw=3,label="momentum_all")
            ax6.plot(x,y2,color='r',lw=3,label="momentum_cat")
            plt.legend()

            ax7 = plt.subplot2grid((25, 2), (4, 0), colspan=4,rowspan=1)
            plt.grid(True) 
            
            y1 = xls_df_day["日K值"][start_day:end_day]      
            y4 = xls_df_day["日D值"][start_day:end_day]      
                
            ax7.plot(x,y1,color='blue',lable = "K-value")                    
            ax7.plot(x,y4,color='r',lable = "D-value") 
            ax7.yaxis.set_label_position("left")
            qualified_day(start_day,end_day,xls_df_day_1)
            plt.legend()                
