# -*- coding: utf-8 -*-
"""
Created on Fri Mar 16 00:27:21 2018

@author: YAO
"""

from datetime import date
import pandas as pd
import numpy as np
import talib
import math

###個股若交易不活絡，該周會有無收盤價情形，因此刪除
def week_format(week_table):
    for j in range(len(week_table)):
        if week_table["周收盤價"][j] == "" :
            week_table.drop([j],axis = 0, inplace = True)
    week_table.index = range(len(week_table)) 
    return week_table

###確認輸入日期的周數
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
    week = date(int(year),int(month),int(day)).isocalendar()
    return week

###將檔案順序以日期前~後排到現在(上~下) 
def normal(file):
    reverse_data = file.sort_values(by = "日期",ascending = True) #以日期最遠到最近排列
    reverse_data.index = range(len(reverse_data))
    return reverse_data 
            
###加入日均線
def day_moving_average(normal_file):
    close_price = normal_file["日收盤價"] 
    close_price = close_price.astype(float)      
       
    ma5_day_list = close_price.rolling(window=5).mean()
    ma20_day_list = close_price.rolling(window=20).mean()
    ma60_day_list = close_price.rolling(window=60).mean()
    ma10_day_list = close_price.rolling(window=10).mean()        
    ma5_day_list.name = "5日均線"
    ma20_day_list.name = "20日均線"
    ma60_day_list.name = "60日均線"
    ma10_day_list.name = "10日均線"
    if ma5_day_list.dtype ==  "O" :
        ma5_day_list = ma5_day_list.astype(float)
        ma20_day_list = ma20_day_list.astype(float)
        ma60_day_list = ma60_day_list.astype(float)
        ma10_day_list = ma10_day_list.astype(float) 
            
    ema20_day = talib.EMA(np.array(close_price),20) 
    ema20_day_columns = ["日EMA20"]
    ema20_day = pd.DataFrame(ema20_day ,columns = ema20_day_columns)

    day_ma = pd.concat([round(ma5_day_list,2),round(ma10_day_list,2),round(ma20_day_list,2),round(ma60_day_list,2),round(ema20_day,2)],axis = 1)
    return day_ma

###加入日漲跌幅
def day_down_up_percentage(normal_file):
    close_price = normal_file["日收盤價"]  
    price_list = [0]
    for i in range(1,len(normal_file)) :
        try:
            price_list.append(round(((close_price[i]-close_price[i-1])/close_price[i-1])*100,2))
        except:
            continue
        
    columns = ["漲跌幅"]
    price_df = pd.DataFrame(price_list,columns = columns)     
    return price_df
    
###加入日乖離率
def day_ma20_ma60_bias(normal_file,day_moving_average):
    
    ma20_bias_list = []
    for i in range(len(day_moving_average["20日均線"])):
        ma20_bias = round((normal_file["日收盤價"][i] - day_moving_average["20日均線"][i]) / day_moving_average["20日均線"][i]*100*20,2)
        if math.isnan(float(ma20_bias)) is False:            
            ma20_bias_list.append(str(ma20_bias))
        else:
            ma20_bias_list.append("")
          
    columns_20 = ["日20均乖離"]
    ma20_bias= pd.DataFrame(ma20_bias_list,columns = columns_20)     
    
    ma60_bias_list = []    
    for i in range(len(day_moving_average["60日均線"])):
        ma60_bias = round((normal_file["日收盤價"][i] - day_moving_average["60日均線"][i]) / day_moving_average["60日均線"][i]*100*20,2)
        if math.isnan(float(ma60_bias)) is False:            
            ma60_bias_list.append(str(ma60_bias))
        else:
            ma60_bias_list.append("")

    columns_60 = ["日60均乖離"]
    ma60_bias= pd.DataFrame(ma60_bias_list,columns = columns_60) 

    ma5_bias_list = [] 
    for i in range(len(day_moving_average["5日均線"])):
        ma5_bias = round((normal_file["日收盤價"][i] - day_moving_average["5日均線"][i]) / day_moving_average["5日均線"][i]*100*20,2)
        if math.isnan(float(ma5_bias)) is False:            
            ma5_bias_list.append(str(ma5_bias))
        else:
            ma5_bias_list.append("")

    columns_5 = ["日5均乖離"]
    ma5_bias= pd.DataFrame(ma5_bias_list,columns = columns_5) 
      
    day_bias_day = pd.concat([ma5_bias,ma20_bias,ma60_bias],axis = 1) 
    
    return day_bias_day

###加入日均量    
def moving_average_volume(normal_file):
    volume = normal_file["成交股數"]
    ma5_volume_list = volume.rolling(window=5).mean()
    ma5_volume_list.name = "5日均量"
    day_vo = pd.concat([ma5_volume_list],axis = 1) #,ma10_volume_list,ma20_volume_list
    return day_vo

###加入日k MACD
def day_macd(normal_file):
    close_price = normal_file["日收盤價"]
    close_price = close_price.astype(float)
    macd, signal, hist = talib.MACD(close_price.values, fastperiod=12, slowperiod=26, signalperiod=9)  
    macd_list=[]
    signal_list = []
    hist_list = []
    for num in range(len(macd)):
        macd_list.append(round(macd[num]*100,2))
        signal_list.append(round(signal[num]*100,2))
        hist_list.append(round(hist[num]*100,2))   
            
    columns = ["日MACD"]
    MACD_array = pd.DataFrame(macd_list,columns = columns) 
    columns1 = ["日signal"]
    signal_array = pd.DataFrame(signal_list,columns = columns1) 
    columns2 = ["日MACD-OSC"]
    MACD_OSC_array = pd.DataFrame(hist_list,columns = columns2)
    day_macd = pd.concat([MACD_array,signal_array,MACD_OSC_array],axis = 1)
    return day_macd

###加入日k KD
def day_kd(normal_file,number):
    ###計算RSV
    close_price = normal_file["日收盤價"].astype(float)
    low_price = normal_file["日最低價"].astype(float)
    high_price = normal_file["日最高價"].astype(float)
    RSV_kd = round(100* (( close_price - low_price.rolling(window=9).min() ) / (high_price.rolling(window=9).max() - low_price.rolling(window=9).min())),2)
    RSV = RSV_kd.replace("NaN",0)
    RSV.name = "日RSV"
    
    xls_file_kd = pd.ExcelFile(r"D:\\Stock Investment\STOCK-KD.xls")
    xls_df_kd = xls_file_kd.parse(xls_file_kd.sheet_names[0])
    ###開啟kd檔案由裡面讀取起始kd日期，搜尋股價編號如果和上市檔案一致，記錄該股kd，並把kd起始日期記為kd_date
    for i in range(len(xls_df_kd)):
        if xls_df_kd["股票代號"][i] == int(number):
            k_value = float(xls_df_kd["日K(9)"][i])
            d_value = float(xls_df_kd["日D(9)"][i])
            kd_date = int(xls_df_kd["日期"][i])
            break   
    ###以日期長度為主，如果日期等於kd的起始日期，則填入kd值
    empty_list_k = []
    empty_list_d = []
    for j in range(len(normal_file["日期"])): 
        if int(normal_file["日期"][j]) == kd_date :
            empty_list_k.append(k_value)
            empty_list_d.append(d_value)
            date_locate = j     #紀錄起始kd為第幾格，要算之後日期的kd值
        else:                   #非起始日期的日子，則填入空白格
            empty_list_k.append("") 
            empty_list_d.append("")
            
    columns1 = ["日K值"]
    K_df = pd.DataFrame(empty_list_k,columns = columns1)  
    columns1 = ["日D值"]
    D_df = pd.DataFrame(empty_list_d,columns = columns1)

    
    day_kd =pd.concat([RSV, K_df, D_df],axis = 1)
    ###計算kd起始值後kd值
    for i in range(date_locate+1,len(normal_file["日期"])):
        day_kd.iloc[i,1] = round((2/3)*day_kd.iloc[i-1,1]+(1/3)*day_kd.iloc[i,0],2)
        day_kd.iloc[i,2] = round((2/3)*day_kd.iloc[i-1,2]+(1/3)*day_kd.iloc[i,1],2)
   
    return day_kd

###加入日k 布林
def day_bbands(normal_file,day_moving_average,cycle_day,number):
    close_price = normal_file["日收盤價"]
    close_price = np.array(close_price.astype(float))  # 有些檔案normal_file["日收盤價"]會是object，因此弄成float64
    upper, middle, lower = talib.BBANDS(close_price, timeperiod=20, nbdevup=2, nbdevdn=2, matype=0)
            
    data = []
    data_all = []
    for num in range(len(upper)):
       data.append(round(upper[num],2))
       data.append(round(middle[num],2))
       data.append(round(lower[num],2))
       data_all.append(data)
       data = []

    column_title_formonth = ["日上軌","日中軸","日下軌"]
    day_bbands = pd.DataFrame(data_all,columns = column_title_formonth )

    return day_bbands

###加入日k %B-收盤
def day_perb(normal_file,day_bbands):
    data = []
    percent_b_list = []
    for i in range(len(normal_file["日收盤價"])):
        b_value = (normal_file["日收盤價"][i]-day_bbands["日下軌"][i])/(day_bbands["日上軌"][i]-day_bbands["日下軌"][i])
        data.append(round(b_value*100,2))            
        percent_b_list.append(data)  
        data = []   

    column_title_formonth = ["日%B"]
    day_perb = pd.DataFrame(percent_b_list,columns = column_title_formonth )
    
    return day_perb 
 
###加入日k 帶寬指標
def day_band_width(day_bbands):
    data = []
    data_all = []
    for i in range(len(day_bbands)):
        band_width = (day_bbands["日上軌"][i]-day_bbands["日下軌"][i])/day_bbands["日中軸"][i]
        data.append(round(band_width,4)*100)
        data_all.append(data)
        data = []
        
    column_title_formonth = ["日帶寬指標"]
    day_band_width = pd.DataFrame(data_all,columns = column_title_formonth )
                
    return day_band_width 

def pattern(normal_file):

    std_y1,nor_y1,mu_y1 = [],[],[]    
    
    y1 = normal_file["日收盤價"]            
    for k in range(len(y1)):
        if k < 59:
            std_y1.append("")
            mu_y1.append("")                    
        else:
            std_y1.append(np.std(y1[k-59:k+1]))
            mu_y1.append(np.mean(y1[k-59:k+1]))

    for k in range(len(mu_y1)):
        if k < 59:
            nor_y1.append("")
        else:
            nor_y1.append((y1[k]-mu_y1[k])/std_y1[k])   
            
    nor_y1 = pd.DataFrame(nor_y1 ,columns = ["型態(季)"])  
    
    return nor_y1

###算出周K##
###邏輯為i=0/1/2時，還不會用到try，但是已經把資訊位置放進week_list = []，接下來下週的(i=3/4/5/6/7)進行時，就會跑到try，close_last_week以及open_last_week
###將會被上周的資訊位置來決定，同時(i=3/4/5/6/7)就會放進清空的week_list，再利用上周已放進weeklist裡面的資訊去做下半部close_price的code
def week_k(normal_file):

    week_info = []          #周的資訊放在裡面(開盤/最高/最低/收盤)
    week_data_array=[]      #創造array把week_info = []放進去
    week_info_locate = []   #紀錄把資料放在哪個位置，周K資訊要放在對應的地方
    week_list = []          #將同周數資料的放在裡面
    close_price_list = []   #每周收盤價
    excel_date = normal_file["日期"]            
    start_week = 0
    close_last_week = 0
    open_last_week = 0      
        
    for j in range(len(normal_file)):
        data_date = str(int(excel_date[j]))
        
        if be_format(data_date)[1] == start_week:
            week_list.append(j) 
        else:
            try:
                close_last_week = week_list[-1]         #上周最後一天
                open_last_week = week_list[0]           #上周第一天
                start_week = be_format(data_date)[1]      
                week_list=[]
                week_list.append(j)
                week_info_locate.append(close_last_week)  #各周星期五的位置，將資料要放進的位置放入此list
            except:
                start_week = be_format(data_date)[1]      
                week_list=[]
                week_list.append(j)
            close_price = normal_file.iloc[close_last_week:close_last_week+1,7:8] 
            exact_close_price = str(close_price)[-5:].strip()
            open_price = normal_file.iloc[open_last_week:open_last_week+1,4:5]
            exact_open_price = str(open_price)[-5:].strip()
            highest = normal_file.iloc[open_last_week: close_last_week+1,5:6].max()
            highest_price = str(highest)[4:str(highest).find("\n")].strip()
            lowest = normal_file.iloc[open_last_week: close_last_week+1,6:7].min()
            lowest_price = str(lowest)[4:str(lowest).find("\n")].strip()
            week_info.append(exact_open_price)
            week_info.append(highest_price)
            week_info.append(lowest_price) 
            week_info.append(exact_close_price)
            week_data_array.append(week_info)
            week_data_pure = week_data_array[1:]       #把i=0/1/2的資訊拿掉,各周的資訊
            close_price_list.append(exact_close_price)
            week_close_price = close_price_list[1:]     #把i=0/1/2的資訊拿掉,各周的收盤價，下面算均線用
            week_locate = week_info_locate        #讓周MA計算用
            week_info = []
        
        if j == len(normal_file)-1:
            close_last_week = week_list[-1]         #上周最後一天
            open_last_week = week_list[0]           #上周第一天    
            week_info_locate.append(close_last_week)  #各周星期五的位置，將資料要放進的位置放入此list
            close_price = normal_file.iloc[close_last_week:close_last_week+1,7:8] 
            exact_close_price = str(close_price)[-5:].strip()
            open_price = normal_file.iloc[open_last_week:open_last_week+1,4:5]
            exact_open_price = str(open_price)[-5:].strip()
            highest = normal_file.iloc[open_last_week: close_last_week+1,5:6].max()
            highest_price = str(highest)[4:str(highest).find("\n")].strip()
            lowest = normal_file.iloc[open_last_week: close_last_week+1,6:7].min()
            lowest_price = str(lowest)[4:str(lowest).find("\n")].strip()
            week_info.append(exact_open_price)
            week_info.append(highest_price)
            week_info.append(lowest_price) 
            week_info.append(exact_close_price)
            week_data_array.append(week_info)
            week_data_pure = week_data_array[1:]       #把i=0/1/2的資訊拿掉,各周的資訊
            close_price_list.append(exact_close_price)
            week_close_price = close_price_list[1:]     #把i=0/1/2的資訊拿掉,各周的收盤價，下面算均線用
            week_locate = week_info_locate        #讓周MA計算用
            week_info = []           
    list_data = []
    list_all = []
    for data in week_data_pure:
        for number in data:
            list_data.append(float(number))
        list_all.append(list_data)
        list_data = []
    week_float = list_all

    column_title_forweek = ["周開盤價","周最高價","周最低價","周收盤價"]
    week_k = pd.DataFrame(week_float,columns = column_title_forweek )

           
    return week_k,week_data_pure,week_close_price,week_locate,week_float

###增加各週均線#
def week_moving_average(normal_file,week_float,week_locate,week_close_price):
    new = [] 
    for i in week_close_price:
        new.append(float(i))
    week_close_price = new
    ###增加周5日均線###               
    ma5_week_list = []
    ma5 = 0
    for l in range(len(week_close_price)-4):
        ma5 = (float(week_close_price[l])+float(week_close_price[l+1])+float(week_close_price[l+2])+float(week_close_price[l+3])+float(week_close_price[l+4]))/5
        ma5_week_list.append(round(ma5,3))
    for num1 in range(len(ma5_week_list)):
        week_float[num1+4].append(ma5_week_list[num1])          

    ###增加周10日均線###
    ma10_week_list = []
    ma10 = 0

    for m in range(len(week_close_price)-9):
        ma10 = (float(week_close_price[m])+float(week_close_price[m+1])+float(week_close_price[m+2])+float(week_close_price[m+3])+float(week_close_price[m+4])+
                float(week_close_price[m+5])+float(week_close_price[m+6])+float(week_close_price[m+7])+float(week_close_price[m+8])+float(week_close_price[m+9]))/10
        ma10_week_list.append(round(ma10,2))

    for num1 in range(len(ma10_week_list)):
        week_float[num1+9].append(ma10_week_list[num1])
        
    ###增加周月均線###
    ma20_week_list = []
    ma20 = 0

    for m in range(len(week_close_price)-19):
        ma20 = (float(week_close_price[m])+float(week_close_price[m+1])+float(week_close_price[m+2])+float(week_close_price[m+3])+float(week_close_price[m+4])+
                float(week_close_price[m+5])+float(week_close_price[m+6])+float(week_close_price[m+7])+float(week_close_price[m+8])+float(week_close_price[m+9])+float(week_close_price[m+10])+
                float(week_close_price[m+11])+float(week_close_price[m+12])+float(week_close_price[m+13])+float(week_close_price[m+14])+float(week_close_price[m+15])+float(week_close_price[m+16])+
                float(week_close_price[m+17])+float(week_close_price[m+18])+float(week_close_price[m+19]))/20
        ma20_week_list.append(round(ma20,2))

    for num1 in range(len(ma20_week_list)):
        week_float[num1+19].append(ma20_week_list[num1])    
        
    week_empty_array= []
    week_empty=["","","","","","",""]
        
    for o in range(len(normal_file["日期"])):
        week_empty_array.append(week_empty)
    for locate_number in range(len(week_locate)):
        week_empty_array[week_locate[locate_number]]=week_float[locate_number]
    column_title_forweek = ["周開盤價","周最高價","周最低價","周收盤價","周5日均線","周10日均線","周20日均線"]
    week_ma = pd.DataFrame(week_empty_array,columns = column_title_forweek )

    return week_ma

###加入周乖離率
def week_ma20_ma5_bias(normal_file,week_float,week_locate):
    column_title_forweek = ["周開盤價","周最高價","周最低價","周收盤價","周5日均線","周10日均線","周20日均線"]
    week_ma = pd.DataFrame(week_float,columns = column_title_forweek ) 
        
    ma5_bias_list = []
    for i in range(len(week_ma["周收盤價"])):
        try:
            ma5_bias = round(((float(week_ma["周收盤價"][i])) - float(week_ma["周5日均線"][i])) / float(week_ma["周5日均線"][i])*100,2)
            ma5_bias_list.append(str(ma5_bias))
        except:
            ma5_bias_list.append("")

    ma20_bias_list = []
    for i in range(len(week_ma["周收盤價"])):
        try:         
            ma20_bias = round(((float(week_ma["周收盤價"][i])) - float(week_ma["周20日均線"][i])) / float(week_ma["周20日均線"][i])*100,2)
            ma20_bias_list.append(str(ma20_bias))
        except:
            ma20_bias_list.append("")
       
    list_data = []
    list_all = []
    for i in range(len(ma20_bias_list)):
        list_data.append(ma5_bias_list[i])
        list_data.append(ma20_bias_list[i])
        list_all.append(list_data)
        list_data = []
    ma5_ma20_list = list_all

    week_empty_array= []
    week_empty=["",""]
        
    for o in range(len(normal_file["日期"])):
        week_empty_array.append(week_empty)

    for locate_number in range(len(week_locate)):
        week_empty_array[week_locate[locate_number]]=ma5_ma20_list[locate_number]

    columns = ["周5均乖離","周20均乖離"]
    ma5_ma20_df= pd.DataFrame(week_empty_array,columns = columns) 

    return ma5_ma20_df
    
###加入周k MACD
def week_macd(normal_file,week_float,week_locate):
    column_title_forweek = ["周開盤價","周最高價","周最低價","周收盤價","周5日均線","周10日均線","周20日均線"]
    week_data = pd.DataFrame(week_float,columns = column_title_forweek ) 

    close_week_price = week_data['周收盤價'].astype(float)
    macd, signal, hist = talib.MACD(close_week_price.values, fastperiod=12, slowperiod=26, signalperiod=9)  
    #弄成list型態後,才能夠與空list調換，符合格式        
    macd_list = list(macd)
    signal_list = list(signal)
    hist_list = list(hist)
    list_data = []
    list_all = []
    for number in range(len(macd)):
       list_data.append(round(macd_list[number]*100,2))
       list_data.append(round(signal_list[number]*100,2))
       list_data.append(round(hist_list[number]*100,2))
       list_all.append(list_data)
       list_data = []            
       
    week_empty_array= []
    week_empty=["","",""]
    for o in range(len(normal_file["日期"])):
        week_empty_array.append(week_empty)

    for locate_number in range(len(week_locate)):
        week_empty_array[week_locate[locate_number]]=list_all[locate_number]
        
    column_title_forweek = ["周MACD","周signal","周MACD-OSC"]
    week_macd = pd.DataFrame(week_empty_array,columns = column_title_forweek )
    return week_macd        
        
###加入周k KD
def week_kd(normal_file,week_k,week_locate,number):

    ###計算RSV
    close_price = week_k["周收盤價"].astype(float)
    low_price = week_k["周最低價"].astype(float)
    high_price = week_k["周最高價"].astype(float)
    RSV_kd = round(100* (( close_price - low_price.rolling(window=9).min() ) / (high_price.rolling(window=9).max() - low_price.rolling(window=9).min())),2)
    RSV = RSV_kd.replace("NaN",0)
    RSV.name = "周RSV"

    xls_file_kd = pd.ExcelFile("D:\\Stock Investment\\STOCK-KD.xls")
    xls_df_kd = xls_file_kd.parse(xls_file_kd.sheet_names[0])
    for i in range(len(xls_df_kd)):
        if xls_df_kd["股票代號"][i] == int(number):
            k_value = float(xls_df_kd["週K(9)"][i])
            d_value = float(xls_df_kd["週D(9)"][i])
            kd_date = int(xls_df_kd["日期"][i])
            break
    ###先找日k日期位在哪一欄，去找周k表格看他欄數是第幾個，以此得知第幾個要放起始周k          
    for k in range(len(normal_file["日期"])):
        if kd_date == int(normal_file["日期"][k]):
            locate_kd = week_locate.index(k)
            break

    empty_list_k = []
    empty_list_d = []
    for j in range(len(week_k)): 
        if j == locate_kd :
            empty_list_k.append(k_value)
            empty_list_d.append(d_value)
        else:
            empty_list_k.append("") 
            empty_list_d.append("")

    columns1 = ["周K值"]
    K_df = pd.DataFrame(empty_list_k,columns = columns1)  

    columns1 = ["周D值"]
    D_df = pd.DataFrame(empty_list_d,columns = columns1)

    week_kd =pd.concat([RSV, K_df, D_df],axis = 1)
    
    for i in range(locate_kd+1,len(week_k["周收盤價"])):
        week_kd.iloc[i,1] = round((2/3)*week_kd.iloc[i-1,1]+(1/3)*week_kd.iloc[i,0],2)
        week_kd.iloc[i,2] = round((2/3)*week_kd.iloc[i-1,2]+(1/3)*week_kd.iloc[i,1],2)
        
    data = []
    data_all = []
    for num in range(len(week_kd)):
        data.append(week_kd["周RSV"][num])
        data.append(week_kd["周K值"][num])
        data.append(week_kd["周D值"][num])
        data_all.append(data)
        data = []

        week_empty_array= []
        week_empty=["","",""]        
    for o in range(len(normal_file["日期"])):
        week_empty_array.append(week_empty)
    for locate_number in range(len(week_locate)):
        week_empty_array[week_locate[locate_number]]=data_all[locate_number]

    columns1 = ["周RSV","周K值","周D值"]
    week_kd = pd.DataFrame(week_empty_array,columns = columns1)        
    return week_kd 

def week_bbands(normal_file,week_k_add_day,week_locate,week_moving_average):      
    new_close_prise = []
    for data in week_k_add_day["周收盤價"]:
        new_close_prise.append(float(data))
            
    upper, middle, lower = talib.BBANDS(np.array(new_close_prise), timeperiod=20, nbdevup=2, nbdevdn=2, matype=0)   
            
    data = []
    data_all = []
    for num in range(len(upper)):
       data.append(round(upper[num],2))
       data.append(round(middle[num],2))
       data.append(round(lower[num],2))
       data_all.append(data)
       data = []
    
    week_empty_array= []
    week_empty=["","",""]        
    for o in range(len(normal_file["日期"])):
        week_empty_array.append(week_empty)
    for locate_number in range(len(week_locate)):
        week_empty_array[week_locate[locate_number]]=data_all[locate_number]

    column_title_formonth = ["周上軌","周中軸","周下軌"]
    week_bbands = pd.DataFrame(week_empty_array,columns = column_title_formonth )

    week_bbands_origin = data_all    
    return week_bbands,week_bbands_origin 

###加入周k %B
def week_perb(normal_file,week_k_add_day,week_locate,week_bbands_origin):
    column_title_formonth = ["周上軌","周中軸","周下軌"]
    orgin_table = pd.DataFrame(week_bbands_origin,columns = column_title_formonth )
        
    data = []
    percent_b_list = []
    for i in range(len(week_k_add_day["周收盤價"])):
        b_value = (float(week_k_add_day["周收盤價"][i])-orgin_table["周下軌"][i])/(orgin_table["周上軌"][i]-orgin_table["周下軌"][i])
        data.append(round(b_value*100,2))
        percent_b_list.append(data)  
        data = []

    week_empty_array= []
    week_empty=[""]        
    for o in range(len(normal_file["日期"])):
        week_empty_array.append(week_empty)
    for locate_number in range(len(week_locate)):
        week_empty_array[week_locate[locate_number]]=percent_b_list[locate_number]
    column_title_formonth = ["周%B(季)"]
    week_perb = pd.DataFrame(week_empty_array,columns = column_title_formonth )
    return week_perb 

def week_perb_lowest(normal_file,week_k_add_day,week_locate,week_bbands_origin):
    column_title_formonth = ["周上軌","周中軸","周下軌"]
    orgin_table = pd.DataFrame(week_bbands_origin,columns = column_title_formonth )
        
    data = []
    percent_b_list = []
    for i in range(len(week_k_add_day["周最低價"])):
        b_value = (float(week_k_add_day["周最低價"][i])-orgin_table["周下軌"][i])/(orgin_table["周上軌"][i]-orgin_table["周下軌"][i])
        data.append(round(b_value*100,2))
        percent_b_list.append(data)  
        data = []

    week_empty_array= []
    week_empty=[""]        
    for o in range(len(normal_file["日期"])):
        week_empty_array.append(week_empty)
    for locate_number in range(len(week_locate)):
        week_empty_array[week_locate[locate_number]]=percent_b_list[locate_number]
    column_title_formonth = ["周%B-最低"]
    week_perb = pd.DataFrame(week_empty_array,columns = column_title_formonth )
    return week_perb        
###加入周k 帶寬指標
def week_band_width(normal_file,week_locate,week_bbands_origin):
    column_title_formonth = ["周上軌","周中軸","周下軌"]
    orgin_table = pd.DataFrame(week_bbands_origin,columns = column_title_formonth )
    
    data = []
    data_all = []
    for i in range(len(week_bbands_origin)):
        band_width = (float(orgin_table["周上軌"][i])-float(orgin_table["周下軌"][i]))/float(orgin_table["周中軸"][i])
        data.append(round(band_width,4))
        data_all.append(data)
        data = []

    week_empty_array= []
    week_empty=[""]        
    for o in range(len(normal_file["日期"])):
        week_empty_array.append(week_empty)
    for locate_number in range(len(week_locate)):
        week_empty_array[week_locate[locate_number]]=data_all[locate_number]

    column_title_formonth = ["周帶寬指標"]
    week_band_width = pd.DataFrame(week_empty_array,columns = column_title_formonth )
    return week_band_width

def week_volume_money_sum(week_table):  
    trans_sum,stock_sum,money_sum = 0,0,0
    for i in range(len(week_table)):
        if week_table["周最高價"][i] == "" :
            stock_sum = week_table["成交股數"][i] + stock_sum
            money_sum = week_table["成交金額"][i] + money_sum
            trans_sum = week_table["成交筆數"][i] + trans_sum
        else:
            stock_sum = week_table["成交股數"][i] + stock_sum
            money_sum = week_table["成交金額"][i] + money_sum
            trans_sum = week_table["成交筆數"][i] + trans_sum

            week_table["成交股數"][i] = stock_sum
            week_table["成交金額"][i] = money_sum
            week_table["成交筆數"][i] = trans_sum
     
            trans_sum,stock_sum,money_sum = 0,0,0
    
    return week_table
