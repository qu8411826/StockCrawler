# -*- coding: utf-8 -*-
"""
Spyder Editor
"""
import requests
import pandas as pd
import time
              
IPO_Date = {
    1817:(2013, 10, 24),
    3558:(2013, 12, 30),
    2228:(2013, 11, 25),
    4163:(2012, 11, 15),
    4999:(2013,  6,  3),
    6412:(2013, 11,  8),
    5289:(2013, 11, 27),
    4126:(2013,  1,  1), # this company has no data for 2012, but I don't know why???
    4205:(2013,  1,  1) # this company has no data for 2012, but I don't know why???
}


# 個別現金流量表 before IFRS http://mops.twse.com.tw/mops/web/t05st36
# 合併財務報表 before IFRS http://mops.twse.com.tw/mops/web/t05st33
# 個別現金流量表 after IFRS http://mops.twse.com.tw/server-java/t164sb05
# 合併財務報表 before IFRS http://mops.twse.com.tw/server-java/t164sb01

# get indivisual 合併財務報表預覽
def individual_quarterly_asset_report(type_k, co_id, co_name, year, quarter):

    return

    if (co_id in IPO_Date.keys()) : # this stock is still young
        if (year < IPO_Date[co_id][0] or (year == IPO_Date[co_id][0] and quarter <= ((IPO_Date[co_id][1]/3)+1))) :
            return None # we're looking up material before it's available time

    print("Getting 合併財務報表 for %d %s %dQ%d" %(co_id, co_name, year, quarter))
    # 假如是西元，轉成民國
    if year > 1900: cn_year = year - 1911
    else : cn_year = year

    if (year <= 2012) :
        url = 'http://mops.twse.com.tw/mops/web/t05st33' # before IFRS
        r = requests.post(url, {
                'step':1,
                'firstin':1,
                'off':1,
                'co_id':co_id,
                'year':cn_year,
                'season':'0'+str(quarter)
            })
        r.encoding = 'utf8'
        try:
            dfs = pd.read_html(r.text)
        except:
            print ("[Stock Crawler]: can't get web data from %s" %(url))
            pass # normally, the page will say: 查無資料
        time.sleep(2)

        if (len(dfs[0].index[(dfs[0][0] == '會計科目')]) <= 0) :
            print ("[Stock Crawler]: can't find 會計科目")
            # encodeURIComponent=1&step=1&firstin=1&off=1&keyword4=&code1=&TYPEK2=&checkbtn=&queryName=co_id&TYPEK=sii&isnew=false&co_id=2892&year=101&season=01&check2858=Y
            r = requests.post(url, {
                'encodeURIComponent':1,
                'step':1,
                'firstin':1,
                'off':1,
                'keyword4':'',
                'code1':'',
                'TYPEK2':'',
                'checkbtn':'',
                'queryName':'co_id',
                'TYPEK':'sii',
                'isnew':'false',
                'co_id':co_id,
                'year':cn_year,
                'season':'0'+str(quarter),
                'check2858':'Y' })
            r.encoding = 'utf8'
            try:
                print ("[Stock Crawler]: can't get web data from %s" %(url))
                dfs = pd.read_html(r.text)
            except:

                pass # normally, the page will say: 查無資料
                time.sleep(2)

        
        # get the tables that are meaningful
        # df = pd.concat([df_now for df_now in dfs if (df_now.shape[1] <= 50 and df_now.shape[0] > 3)])
        df = dfs[0]
        # remove invalid rows
        df = df[df.notnull()]        

        if (len(df.index[(df[0] == '會計科目')]) > 0) :
            row_idx_of_columns = df.index[(df[0] == '會計科目')][0]
            df.columns = df.iloc[row_idx_of_columns]
            # move header row to the row next to '會計科目'
            df = df[(row_idx_of_columns+1) :]
            # change the index to '會計科目'
            df = df.set_index('會計科目')    
            # drop last row
            df = df.drop(df.index[df.index.size-1])
            # drop empty columns
            df = df.dropna(axis=1, how='all')
        else :
            print ("[Stock Crawler]: still can't find 會計科目")
            if (co_id in IPO_Date.keys()) : # this stock is still young
                if (year <= IPO_Date[co_id][0] or quarter <= ((IPO_Date[co_id][1]/3)+1)) :
                    return None # we're looking up material before it's available time
            assert(0) # why are we here?
    else :
        url = 'http://mops.twse.com.tw/server-java/t164sb01' # after IFRS           
        r = requests.post(url, {
                'encodeURIComponent':1,
                'step':1,
                'CO_ID':co_id,
                'SYEAR':cn_year,
                'SSEASON':quarter,
                'REPORT_ID':'C'
            })
        r.encoding = 'big5'
        try:
            dfs = pd.read_html(r.text)
        except:
            # normally, the page will say: 查無資料
            time.sleep(2)
            r = requests.post(url, {
                    'encodeURIComponent':1,
                    'step':1,
                    'CO_ID':co_id,
                    'SYEAR':cn_year,
                    'SSEASON':quarter,
                    'REPORT_ID':'A' # different REPORT_ID
                })
            r.encoding = 'big5'
        try:
            dfs = pd.read_html(r.text)
        except:
            return None
        
        if (len(dfs) == 1 and len(dfs[0]) == 1) : return None
        
        # get the tables that are meaningful
        # df = pd.concat([df_now for df_now in dfs if (df_now.shape[1] <= 50 and df_now.shape[0] > 3)])
        df = dfs[1]

        # remove invalid rows
        df = df[df.notnull()]

        if (len(df.index[(df[0] == '會計項目')]) > 0) :
            row_idx_of_columns = df.index[(df[0] == '會計項目')][0]
            df.columns = df.iloc[row_idx_of_columns]
            # move header row to the row next to '會計項目'
            df = df[(row_idx_of_columns+1) :]
            # change the index to '會計項目'
            df.set_index('會計項目')    
            # drop empty columns
            df.dropna(axis=1, how='all')
    return df

