:wave::wave::wave: <br>
您好，我是對資料分析有濃厚興趣的姚廷諺。從 2016 年起，我開始自學 Python，應用於金融資料整理、分析、邏輯設計與系統建置，逐步累積資料清洗、儲存、報表設計與視覺化呈現等經驗。早期資料管理主要以 Excel 建立與維護，並透過 Python 進行資料讀寫與分析，視覺化則以 matplotlib 為主；隨著資料量與分析需求增加，目前調整為以 PostgreSQL 作為主要資料儲存架構，並搭配 Python 進行資料擷取、查詢、整併與分析，視覺化除了 matplotlib 之外，也透過 Grafana dashboard 呈現。<br>
<br>
本專案為本人自行規劃與實作的金融數據分析與推薦系統，目的不僅在於進行資料分析，更在於將資料蒐集、清洗、指標運算、條件篩選、資料儲存與結果視覺化整合為一套可持續執行的系統流程。此專案為我作品集中的個人實作項目之一，主要用來呈現我在系統規劃、資料流程設計、技術整合與實作的能力。
<br>
本專案主要展示我如何運用 Python、Excel、PostgreSQL 與 Grafana 建置資料處理與分析流程，內容包含系統運算邏輯概述，以及系統程式架構與實作說明。文中附件所提供之檔案，主要為早期以 Excel 作為資料管理架構時所保存的市場原始資料，以及透過 Python 計算後產出的結果報表；後續系統調整為以 PostgreSQL 搭配 SQL 指令進行資料處理與結果產出，因此現行架構下較無相同形式之報表檔案可供附上。文中所展示之程式碼並非完整系統全部內容，主要用於說明整體設計思路與相關技術能力。<br>

# 目錄
- [自身能力評估](#自身能力評估)
- [分析工具](#分析工具)
- [專案目標](#專案目標)
- [一、系統運算邏輯概述(資料分析方法)](#一系統運算邏輯概述)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [系統流程圖](#building_construction系統流程圖)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [資料搜集](#spider資料搜集)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [資料清洗](#wrench資料清洗)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [資料庫設計與建置邏輯](#file_folder-資料庫設計與建置邏輯)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [策略運算邏輯](#computer策略運算邏輯)<br>
- [二、程式碼概述(資料分析及系統程式碼)](#二程式碼概述)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [個股報表產生程式碼](#chart_with_upwards_trend個股報表產生程式碼)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [選股系統主要程式碼](#rocket-選股系統主要程式碼)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [可視化圖表程式碼](#bar_chart-可視化圖表程式碼)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [回測策略](#game_die回測策略)<br>
  
<br>
 
# 自身能力評估
1. Python 與資料處理能力：可獨立使用 Python 進行資料擷取、資料清理、資料庫存取、邏輯設計、報表設計與產出等流程，具備從資料蒐集、處理到結果輸出的完整實作經驗。<br>
2. 資料分析與視覺化能力：熟悉使用 Python 進行資料清理、分析與視覺化，並具備 NumPy、Pandas 等資料處理工具實務經驗。早期以 matplotlib 進行圖表呈現，後續隨資料架構調整，亦搭配 Grafana dashboard 進行視覺化展示，以提升資料檢視與應用效率。<br>
3. Excel 與報表製作能力：熟悉 Excel 資料整理、函數應用（如 VLOOKUP）、報表彙整與圖表製作，能依需求將資料整理為清楚易讀之表格、圖表與簡報素材，協助後續溝通與應用；早期資料管理亦主要以 Excel 建立與維護，具備以 Python 操作 Excel 進行資料讀寫、整理與分析經驗。<br>
4. 資料庫操作能力：熟悉 PostgreSQL 資料表設計、資料儲存、查詢與維護，並能透過 Python 連接資料庫進行資料擷取與應用，同時使用 pgAdmin 進行資料庫操作與管理。<br>
5. 系統與後台操作經驗：具網站與系統規劃相關經驗，因工作需協助客戶進行系統設計、功能確認與流程測試，對於系統流程理解、後台操作及功能測試具備實務基礎。<br>
6. 跨部門協作與溝通能力：具七年專案管理經驗，過程中需與專案成員、供應商及客戶進行需求確認、進度追蹤與問題協調，因此具備良好的跨部門合作、溝通協調與執行能力。<br>
7. 邏輯分析、資料品質控管與問題解決能力：能針對資料與實務需求進行邏輯設計、結果驗證與流程優化，並重視資料正確性、一致性與作業時效，持續提升資料應用與報表產出的品質。<br>

# 分析工具
1. 資料分析語言：Python、SQL。<br>
2. 資料儲存架構：早期市場原始資料主要儲存在 Excel，個股量化報表儲存在 CSV；目前調整為以 PostgreSQL 作為主要資料儲存與查詢架構。<br>
3. 主要使用套件：pandas、numpy、talib（技術指標）。<br>
4. 主要視覺化工具：matplotlib、mpl_finance、Grafana dashboard。<br>
5. 資料操作與分析模組：report_func、daily_count_func、index_all（皆為親自撰寫）。<br>

# 專案目標
1. 上市/上櫃股數高達1900檔，該如何從中挑選有潛力的個股?<br>
2. 法人、散戶與市場資金行為皆反映於每日揭露資訊中，如何從大量且分散的資料中整理出可供分析與判讀的邏輯脈絡？<br>

# 一、系統運算邏輯概述

## :building_construction:系統流程圖
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E7%B3%BB%E7%B5%B1%E6%B5%81%E7%A8%8B%E5%9C%96_0415.png)

## :spider:資料搜集
爬蟲上市/上櫃所有個股資料，包含以下項目，系統於交易日 14:35 開始自動執行資料擷取流程<br>
- 每日收盤行情<br>
- 三大法人買賣超日報<br>
- 融資融券餘額<br>
- 融券借券賣出餘額<br>
- 外資及陸資投資持股統計<br>
- 大盤指數及各類股指數與交易金額<br>

(以下為檔案範例，請點擊後下載查看，此為早期以 Excel 管理架構時的原始資料檔案)<br>

●[信錦-每日行情](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%BF%A1%E9%8C%A6-%E6%AF%8F%E6%97%A5%E8%A1%8C%E6%83%85.xls) 
●[信錦-三大法人買賣超](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%BF%A1%E9%8C%A6-%E4%B8%89%E5%A4%A7%E6%B3%95%E4%BA%BA%E8%B2%B7%E8%B3%A3%E8%B6%85.xls)
●[信錦-融資融券彙總](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%BF%A1%E9%8C%A6-%E8%9E%8D%E8%B3%87%E8%9E%8D%E5%88%B8%E5%BD%99%E7%B8%BD.xls)
●[信錦-融券借券彙總](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%BF%A1%E9%8C%A6-%E8%9E%8D%E5%88%B8%E5%80%9F%E5%88%B8%E5%BD%99%E7%B8%BD.xls)
●[信錦-外資及陸資持股](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%BF%A1%E9%8C%A6-%E5%A4%96%E8%B3%87%E5%8F%8A%E9%99%B8%E8%B3%87%E6%8C%81%E8%82%A1.xls)

## :wrench:資料清洗
每個檔案的欄位依照網站為主，主要做以下處理，儲存後檔案畫面如下圖。<br>
1. 刪除不需要的欄位及其資訊<br>
2. 統一日期格式<br>
3. 轉換資料型態(由 str 轉換成 float64 )<br>

Excel管理架構<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/Daily%20Quotes.png)

SQL管理架構<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/postgresql_1101_daily_info.png)

## :file_folder: 資料庫設計與建置邏輯
1. 早期資料管理主要以 Excel 架構進行，依官方上市櫃股票分類建立產業大項，再依個股名稱建立資料夾，將相關資訊分別存放於不同的 Excel 檔案中，如下方 Excel 畫面所示，以每日開高低收資料為例，Excel 架構下需為每一檔股票各自建立名為 "股名-每日行情" 檔案紀錄。<br>
2. 後續調整為 PostgreSQL 架構後，則改以資料表集中管理各項爬蟲資料，每個資料項目僅需維護一張資料表，並將所有股票資料統一存放於表中，再透過 symbol（股票代號）進行查詢即可；同時維護一張所有股票的資料總表方便查詢該股屬於哪個產業以及哪個市場及其他個股基本資訊。<br><br>

Excel管理架構<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/Stock%20Category.png)

SQL管理架構(股票資料總表)<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/postgresql_stock_info.png)

SQL管理架構，使用symbol= '1101' 操作結果<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/postgresql_1101_daily_info.png)



## :computer:策略運算邏輯
當策略運算所需的資料都爬蟲完成之後，便開始進行策略運算，步驟如下：<br>
1. 將所有股票做分類。<br>
2. 利用當日個股資料計算各分類籌碼流向結果並計算籌碼流向指標(下圖圖一中藍色實線，標註為"Evaluation Index")，單一分類計算結果範例請見下方檔案"產業計算紀錄表"與圖一說明。<br>
3. 所有分類計算完成後，將各分類籌碼流向指標分數由高到低做排名。(依據圖一Y座標值做為分數，進行分類總體排名，Y值越高則該群排名越前面)。<br>
4. 計算個股月/周/日三級別報表的所有量化指標。<br>
5. 定義強勢分類的名次範圍，先篩選出強勢分類，再根據所篩選出的分類篩選出符合策略條件的個股，個股計算結果範例請見下方檔案"信錦-day"、"信錦-week"與圖二說明。<br>
6. 產出當日總體推薦個股報表。<br>
<br>

●[產業計算紀錄表](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E6%97%8F%E7%BE%A4%E8%B3%87%E9%87%91.xls)，說明：<br>
1. 檔案中的開始日期(檔案中黃色背景欄位)若有值，即為自定義的籌碼開始進入波動區段，並藉由找出該區段的高低點，得知流進區間內的籌碼流進流出的情況，計算籌碼流向指標。<br>
2. 檔案中的開始日期(檔案中黃色背景欄位)若無值，則尚未進入自定義的籌碼波動區段，因此該族群不會有任何動作，族群內的個股也不會被選上。<br>
3. 運用 matplotlib 將指標以時間序列顯示如圖一。<br>

●[信錦-day](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%BF%A1%E9%8C%A6-day.csv)、[信錦-week](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%BF%A1%E9%8C%A6-week.csv)，說明：<br>
1. 每個個股都有獨立的計算報表，將力道指標及其他判斷指標的計算結果紀錄在各自的報表上，每張表欄位一致。<br>
2. 力道指標(檔案中黃色背景欄位)若小於0，則有機會有一段漲幅，可以隔日開盤立即買進，有機會賺到正報酬。<br>
3. 運用 matplotlib 將指標以及價格以時間序列顯示如圖二。<br>

圖一：圖中示範族群於20240305-20240415、20240515-20240712進入籌碼波動區間，並藉由紅點與綠點定義出籌碼正在流進或流出，再藉以Y的值作為評價分數，將所有的族群以此方式做排名。<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E7%94%A2%E6%A5%AD%E8%A9%95%E5%83%B9%E8%A7%A3%E8%AA%AA.jpg)

圖二：圖中示範個股，藍色粗線為動能指標的連續動態值，可以見到藍色粗線值若低於0(紅色虛線"--"表示)，有機會有一段漲幅，便可以在小於0出現買入訊號時(可見圖中"Buy Signal")，於次交易日買進，有機會賺到正報酬。<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E5%80%8B%E8%82%A1Demo.jpg)
<br>

# 二、程式碼概述
## :chart_with_upwards_trend:個股報表產生程式碼
主要程式碼參考位置：https://github.com/yao790609/Project_Demo_Contract/blob/master/stock_report-listed.py<br>
各項指標運算參考位置(僅列出基本指標程式碼)：https://github.com/yao790609/Project_Demo_Contract/blob/master/index_all.py<br>

說明：<br>
stock_report-listed為對上市股票針對個股資訊產出個別報表，時間頻率為日/周/月三種報表，在此對資料再次做清洗及處理，再引入 index_all 內的 function 做計算，以下情況會對資料做處理：<br>
1. 個股當日無交易但爬蟲有資料。<br>
2. 去除個股資料不需要的欄位。<br>
3. 對於設定區間外的資料做刪除。<br>
   
## :rocket: 選股系統主要程式碼
參考位置(此為程式碼精簡版)：https://github.com/yao790609/Project_Demo_Contract/blob/master/daily_info_update_listed.py<br>

### 此程式碼包含如下功能
 - 資料爬蟲<br>
 - 檢查是否有新上市股票並自動建立資料庫及其資料檔案<br>
 - 計算個股漲跌幅資訊<br>
 - 計算個股日/周/月報表<br>
 - 計算所有分類籌碼流向指標<br>
 - 計算總體選股策略與推薦報表產出<br>

### 系統詳細流程圖
此流程圖紅字地方為系統產出檔案，以下檔案對應下方流程圖編號，若下方沒有則會在後面的流程說明中附上。 <br>
編號1. [市盤](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E5%B8%82%E7%9B%A4.xlsx)<br>
編號4. [漲跌幅報表-上市](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E6%BC%B2%E8%B7%8C%E5%B9%85%E5%A0%B1%E8%A1%A8-%E4%B8%8A%E5%B8%82.xlsx)<br>
編號5. [個股計算開關-市](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E5%80%8B%E8%82%A1%E8%A8%88%E7%AE%97%E9%96%8B%E9%97%9C-%E5%B8%82.xlsx)<br>
編號18. [股票分類-下載用(去除重複)-上市](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E8%82%A1%E7%A5%A8%E5%88%86%E9%A1%9E-%E4%B8%8B%E8%BC%89%E7%94%A8(%E5%8E%BB%E9%99%A4%E9%87%8D%E8%A4%87)-%E4%B8%8A%E5%B8%82.xls)<br>
編號19. [新股清單-市](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E6%96%B0%E8%82%A1%E6%B8%85%E5%96%AE-%E5%B8%82.xlsx)<br>
<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/System%20Flow%20Chart.jpg)

### 在此系統中，載入下列四項自己編寫的Library
daily_count_func / report_func / price_slope_daily / strategy_function<br>

#### daily_count_func
參考位置：https://github.com/yao790609/Project_Demo_Contract/blob/master/index_all.py<br>
作用：目的為提供下方 report_func 的計算函式，為免去每天都要重新產出個股報表耗費大量時間，因此將產出個股報表的指標 Library-index_all 內的 function 調整成根據前一天的結果，接續計算今日個股的日報表與周報表每個欄位的值。<br>

#### report_func
參考位置：https://github.com/yao790609/Project_Demo_Contract/blob/master/report_func.py<br>
作用：引入上方解說的 Library-daily_count_func 內的 function 計算今日指標數值。<br>

#### price_slope_daily
參考位置：https://github.com/yao790609/Project_Demo_Contract/blob/master/price_slope_daily.py<br>
作用：目的為計算每個候選股票在特定的參考指標下，該指標達到特定數值所累積的天數，以此比較個股的累積強度，最後產出總體比較報表。<br>
考參考以下檔案。<br>

●[上市總表](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%B8%8A%E5%B8%82%E7%B8%BD%E8%A1%A8.xls)<br>

#### strategy_function
參考位置(第235行以下)：https://github.com/yao790609/Project_Demo_Contract/blob/master/daily_report.py<br>
作用：目的為將上市總表根據總體篩選策略標準，將符合標準的個股由指標大到小排列，最後產結果報表。<br>
先產出上市結果報表，再產出強勢股列表，最後產出市候選，考參考以下檔案。<br>

●[上市結果報表](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%B8%8A%E5%B8%82%E7%B5%90%E6%9E%9C%E5%A0%B1%E8%A1%A8.xls)<br>
●[強勢股列表-上市](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E5%BC%B7%E5%8B%A2%E8%82%A1%E5%88%97%E8%A1%A8-%E4%B8%8A%E5%B8%82.xls)<br>
●[市候選](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E5%B8%82%E5%80%99%E9%81%B8.xlsx)<br>

## :bar_chart: 可視化圖表程式碼
程式碼參考位置：https://github.com/yao790609/Project_Demo_Contract/blob/master/index_vision.py<br>
1. 由於 matplotlib 的呈現方式以單張靜態圖表為主，較適合用於說明資料分析結果，因此本區域將以 matplotlib 圖表作為說明重點，藉以展示資料經整理、計算與視覺化後的呈現方式，以及各項指標與股價之間的對應關係。<br>
2. 相較之下，Grafana 屬於 web-based dashboard 服務，除可將資料以圖表方式整合呈現外，亦可依需求彈性調整查詢條件，例如股價資訊查詢區間、股票代號及其他相關觀察欄位，便於即時掌握個股狀況與資料變化，因此本頁所附 Grafana 圖面僅作為示意用途，若有面試機會，我可進一步展示 Grafana 的實際操作方式、查詢邏輯與互動效果。<br>
<br>

圖一：光聖在 2024/01/29 以及 2024/03/05 突破布林通道且力道指標小於0，出現買入訊號，可以在次交易日進行購買，有機會賺取正報酬。<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E5%85%89%E8%81%96.jpg)

圖二：冠德在 2023/11/27 以及 2024/03/27 突破布林通道且力道指標小於0，出現買入訊號，可以在次交易日進行購買，有機會賺取正報酬。<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E5%86%A0%E5%BE%B7.jpg)

圖三：正淩在 2024/01/23 以及 2024/05/29 突破布林通道且力道指標小於0，出現買入訊號，可以在次交易日進行購買，有機會賺取正報酬。<br>
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E6%AD%A3%E6%B7%A9.jpg)

圖四： Grafana 視覺化圖表展示
![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/grafana_demo.png)

# :game_die:回測策略
參考位置_1：https://github.com/yao790609/Project_Demo_Contract/blob/master/daily_report.py<br>
參考位置_2：https://github.com/yao790609/Project_Demo_Contract/blob/master/backtesting.py<br>
作用：當個股的日報表與周報表產出之後，以此進行每日的候選股策略回測，我們利用 daily_report.py 的程式碼將候選表產生出來之後(產生報表的邏輯與上述相同)，再利用 backtesting.p y進行回測，此為第一階段的電腦篩選，並無經過個人經驗與技術型態過濾，因此只要候選股在隔天符合買入標準便觸發購買機制，最後產出如下檔案。<br>
2016 年是回測九年來虧損的一年； 2020 年是收益與往年平均差不多的一年，因此放上此兩年的檔案供參考。<br>

●[交易清單-市2016](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%BA%A4%E6%98%93%E6%B8%85%E5%96%AE-%E5%B8%822016.xlsx)<br>
●[交易清單-市2020](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/%E4%BA%A4%E6%98%93%E6%B8%85%E5%96%AE-%E5%B8%822020.xlsx)<br>

說明：由下圖得知， 2016 年全年機器選擇筆數高達 307 筆，損失來到十萬，但經由人為判斷之後，筆數降為 12 筆，損失降至兩萬左右；而 2020 年全年機器選擇筆數高達  407筆，雖整年獲利高達五十五萬，但在資金不夠的情況下並不允許如此執行，增加人為判斷之後，交易筆數降到 65 筆，獲利雖亦降至三十萬左右，但最大同時投入資金僅為二十五萬，是可以負荷的金額，因此在沒有辦法判斷型態的情況下，經由人為經驗判斷的過程還是需要的，雖然利潤可能因此降低，但卻更有效地挑選出有正獲利的筆數。

![image](https://github.com/yao790609/Project_Demo_Contract/blob/master/Demo%20Files/2016%E8%88%872020%E6%94%B6%E7%9B%8A.jpg)

自 2016 年至今，我已在金融資料分析領域深耕超過八年，相信資料蘊藏著無限的價值，而真正的寶藏往往隱藏在細節之中。資料分析不僅是尋找答案的過程，更是透過縝密的邏輯與創意思維，挖掘關鍵資訊、發掘趨勢，進而轉化為可行的策略。未來，我期待持續精進自己的技能，運用資料驅動決策，協助企業在瞬息萬變的市場中找到最佳發展方向，感謝您的查閱，期待能加入貴司成為團隊的一份子。<br>

