# -*- coding: utf-8 -*-
"""
Spyder Editor
"""
import requests
import pandas as pd
import numpy as np
import time
              
bank_stocks = {
'2801':u'彰化銀行','2809':u'京城銀行','2812':u'台中銀行','2820':u'華票','2834':u'臺灣企銀','2836':'高銀',
'2838':u'聯邦銀行','2845':u'遠東商銀','2849':u'安泰銀行','2897':u'王道銀行'}
stock_stocks = {
'2855':u'統一證券','2856':u'元富證券','6005':u'群益證','6024':u'群益期'}
holding_stocks = {
'2880':u'華南金','2881':u'富邦金','2882':u'國泰金','2883':u'開發金','2884':u'玉山金控','2885':'元大金',
'2886':u'兆豐金','2887':u'台新金控','2888':u'新光金','2889':u'國票金控','2890':u'永豐金控','2891':'中信金',
'2892':u'第一金控','5880':u'合庫金'}
special_stocks = {
'1409':u'新纖','1718':u'中纖','2207':u'和泰汽車','2905':u'三商'}

quarterly_web_map = {
u'SII綜合損益彙總表':"http://mops.twse.com.tw/mops/web/ajax_t163sb04",
u'OTC綜合損益彙總表':"http://mops.twse.com.tw/mops/web/ajax_t163sb04",
u'SII資產負債彙總表':"http://mops.twse.com.tw/mops/web/ajax_t163sb05",
u'OTC資產負債彙總表':"http://mops.twse.com.tw/mops/web/ajax_t163sb05",
u'SII營益分析彙總表':"http://mops.twse.com.tw/mops/web/ajax_t163sb06",
u'OTC營益分析彙總表':"http://mops.twse.com.tw/mops/web/ajax_t163sb06"
}

    
def combined_quarterly_statement(year, quarter, table_now):
    # 假如是西元，轉成民國
    cn_year = year
    if year >= 1900: cn_year = year - 1911
    assert(year >= 2013)

    url = quarterly_web_map[table_now]
    param_hash = {
            'encodeURIComponent':1,
            'step':1,
            'firstin':1,
            'off':1,
            'year':cn_year,
            'season':'0'+str(quarter)
    }
    if (table_now[0:3] == 'SII') :
        param_hash['TYPEK'] = 'sii'
    elif (table_now[0:3] == 'OTC') :
        param_hash['TYPEK'] = 'otc'
    # Get html link
    r = requests.post(url, param_hash)
    # Encode the requested content in utf8
    r.encoding = 'utf8'
    
    # loop through all tables and drop header row
    time.sleep(2)
    print("Crawling web for %s%dQ%d" %(table_now, year, quarter))
    try:
        dfs = pd.read_html(r.text)
    except:
        print ('pd.read_html(', r.text, ')')
        assert(0)
        return None
    for i, df_now in enumerate(dfs):
        # check the shape
        dfs[i].shape
        # move header row to 'columns' and skip it
        df_now.columns = df_now.iloc[0]
        dfs[i] = df_now.iloc[1:]
    # concatenate all dataframes with '--' replaced by NaN
    df = pd.concat(dfs).applymap(lambda x: x if x != '--' else np.nan)
    # remove invalid rows    
    df = df[df[u'公司代號'] != u'公司代號']
    df = df[df[u'公司代號'].notnull()]
    # change the index to '公司代號'
    df = df.set_index(u'公司代號')
    return df

