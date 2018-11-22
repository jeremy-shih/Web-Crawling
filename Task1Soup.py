
# coding: utf-8

# In[ ]:


from bs4 import BeautifulSoup
import requests
import pandas as pd
import random
import time
try:                                    # Support Python 2 and 3
    from urllib import urlopen, build_opener          
except ImportError:                     
    from urllib.request import urlopen, build_opener  
    
# Helper function to concatenate multiple dataframes into one excel sheet
def concat_multiple_dfs(df_list, sheets, file_name, spaces):
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')   
    row = 0
    for dataframe in df_list:
        dataframe.to_excel(writer,sheet_name=sheets,startrow=row , startcol=0)   
        row = row + len(dataframe.index) + spaces 
    writer.save()
    
# Helper function to open url
def open_url(stock, date):
    opener = build_opener()
    opener.addheaders = [('User-agent', 'Mozilla/5.0')]
    stocktable = opener.open('http://www.twse.com.tw/exchangeReport/STOCK_DAY?response=html&date='         + date + '&stockNo=' +str(stock))
    return stocktable.read()

#Extract Data from Months January to May for one stock
def extract_one_stock(stock):
    datelist = ['20180101', '20180201', '20180301', '20180401', '20180501']
    dflist = []
    for date in datelist:        
        page = open_url(stock, date)
        soup = BeautifulSoup(page, 'html.parser')
        table=soup.find('table')
        #Generate lists
        A=[]
        B=[]
        C=[]
        D=[]
        E=[]
        F=[]
        G=[]
        H=[]
        I=[]
        
        for row in table.find("tbody").find_all("tr"):
            cells = row.find_all("td")
            if len(cells) == 9:
                A.append(cells[0].find(text=True))
                B.append(cells[1].find(text=True))
                C.append(cells[2].find(text=True))
                D.append(cells[3].find(text=True))
                E.append(cells[4].find(text=True))
                F.append(cells[5].find(text=True))
                G.append(cells[6].find(text=True))
                H.append(cells[7].find(text=True))
                I.append(cells[8].find(text=True))
            
        df = pd.DataFrame(A,columns=['日期'])
        df['成交股數']=B
        df['成交金額']=C
        df['開盤價']=D
        df['最高價']=E
        df['最低價']=F
        df['收盤價']=G
        df['漲跌價差']=H
        df['成交筆數']=I
        
        dflist.append(df)
    
    empty_headers = []
    for i in range(0,dflist[0].shape[1]):
        empty_headers.append("")

    for i in range(1,len(dflist)):
        dflist[i].columns = empty_headers
    
    concat_multiple_dfs(dflist, str(stock), str(stock)+'.xlsx', 1)

    
# Extract multiple stocks
def extract_mult_stocks(stocks):
    for stock in stocks:
        extract_one_stock(stock)
        wait = random.randint(10,20)
        time.sleep(wait)

# Extract Data for all 52 stocks 
stocks = [1101, 1102, 1216, 1301, 1303, 1326, 1402, 2002, 2105,
          2227, 2301, 2303, 2308, 2317, 2327, 2330, 2352, 2354, 2357, 2382, 2395, 2408,
          2409, 2412, 2454, 2474, 2609, 2610, 2633, 2801, 2823, 2880, 2881, 2882, 2883,
          2884, 2885, 2886, 2887, 2890, 2891, 2892, 2912, 3008, 3045, 3481, 4904, 4938, 5871,
          5880, 6505, 9904]

extract_mult_stocks(stocks)
        
        

