# -*- coding: utf-8 -*-
"""
Spyder Editor
"""
import pandas as pd
import time
  
monthly_web_map = {
'SII每月營業收入彙總表':'http://mops.twse.com.tw/nas/t21/sii/t21sc03_',
'OTC每月營業收入彙總表':'http://mops.twse.com.tw/nas/t21/otc/t21sc03_'
}

def combined_monthly_revenue_report(year, month, table_now):
    global num_web_access
    # 假如是西元，轉成民國
    if year > 1900: cn_year = year - 1911
    else: cn_year = year
    
    url = monthly_web_map[table_now]+str(year)+'_'+str(month)
    if cn_year <= 98:
        url += '.html' # old format
    else:
        url += '_0.html' # new format
    
    # 下載該年月的網站，並用pandas轉換成 dataframe
    print("Crawling web for %s_%d%02d" %(table_now, year, month))
    try:
        html_df = pd.read_html(url)
    except:
        return None
    num_web_access += 1
    if num_web_access % 10 == 9: time.sleep(15)
    
    # 處理一下資料
    if html_df[0].shape[0] > 500:
        df = html_df[0].copy()
    else:
        df = pd.concat([df for df in html_df if df.shape[1] <= 11])
    df = df[list(range(0,10))]
    row_idx_of_columns = df.index[(df[0] == '公司代號')][0]
    df.columns = df.iloc[row_idx_of_columns]
    df['當月營收'] = pd.to_numeric(df['當月營收'], 'coerce')
    df = df[~df['當月營收'].isnull()]
    df = df[df['公司代號'] != '合計']
    # change the index to '公司代號'
    df = df.set_index('公司代號')
    return df

