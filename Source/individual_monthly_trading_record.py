# -*- coding: utf-8 -*-
"""
Spyder Editor
"""
import requests
import pandas as pd
import time

# trading record
def individual_monthly_trading_record(co_id, co_name, year, ref_df):
    date_string = "%4d0101" %(year)
    print("Crawling 月成交資訊 for %d in year %d" %(co_id, year))
    
    if ref_df.loc[co_id, 'TYPE_K'] == u'上櫃' :
        url = 'http://www.tpex.org.tw/web/stock/statistics/monthly/result_st44.php'
        r = requests.post(url, {
                'ajax':'true',
                'l':'zh-tw',
                'yy':year,
                'input_stock_code':co_id
                })
        r.encoding = 'utf-8'
        try:
            dfs = pd.read_html(r.text)
        except:
            return None
        if co_id % 11 == 1: time.sleep(15)

        df = dfs[2]
        assert(df.shape[1] == 9)
        df.columns = df.iloc[0]
        assert(df.columns[0] == u'年')
        assert(df.columns[1] == u'月')
        # change the index to 月
        df = df.set_index(u'月')
    else :
        url = 'http://www.twse.com.tw/exchangeReport/FMSRFK?response=html&date=' + date_string + '&stockNo=' + str(co_id)
        # url = "http://www.twse.com.tw/exchangeReport/STOCK_DAY_AVG?response=html" + '&date=' + date_string + '&stockNo=' + str(co_id)

        # 下載該年月的網站，並用pandas轉換成 dataframe
        try:
            dfs = pd.read_html(url)
        except:
            return None
        if co_id % 11 == 1: time.sleep(15)
        
        df = dfs[0]
        assert(df.shape[1] == 9)
        assert(df.columns[0][1] == u'年度')
        assert(df.columns[1][1] == u'月份')
        df.columns = [df.columns[i][1] for i in range(len(df.columns))]
        # change the index to 日期'
        df = df.set_index(u'月份')
    return df

