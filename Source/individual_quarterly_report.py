# -*- coding: utf-8 -*-
"""
Spyder Editor
"""
import requests
import pandas as pd
import time
from pathlib2 import Path

before_IFRS_url_table = {
    '現金流量表': ('http://mops.twse.com.tw/mops/web/t05st36', '會計項目', 'utf8'),
    '資產負債表': ('http://mops.twse.com.tw/mops/web/t05st33', '營業活動之現金流量-間接法', 'utf8')
}
after_IFRS_url_table = {
    '資產負債表': ('http://mops.twse.com.tw/server-java/t164sb01', '會計項目', 'big5'),
    '綜合損益表': ('http://mops.twse.com.tw/server-java/t164sb01', '會計項目', 'big5'),
    '現金流量表': ('http://mops.twse.com.tw/server-java/t164sb01', '會計項目', 'big5')
}

Report_ID_list = [
    ('合併財報', 'C'),
    ('個別財報', 'A'),
    ('個體財報', 'B')
]

wait_time = 4

def CrawForIFRSreport(url, encoding, co_id, cn_year, quarter):
    param = {
        'encodeURIComponent':1,
        'step':1,
        'CO_ID':co_id,
        'SYEAR':cn_year,
        'SSEASON':quarter,
    }
    for (report_name, report_id) in Report_ID_list:
        dfs = None
        param['REPORT_ID'] = report_id
        while (dfs is None):
            try:
                r = requests.post(url, param)
            except requests.exceptions.ConnectionError:
                # Maybe set up for a retry, or continue in a retry loop
                print ("ConnectionError when getting %s for %d in %d Q%d" %(report_name, co_id, cn_year, quarter))
                print ("Wait %d sec and try again" %(wait_time*10))
                time.sleep(wait_time*10)
                continue
            except requests.exceptions.Timeout:
                # Maybe set up for a retry, or continue in a retry loop
                assert(0)
            except requests.exceptions.TooManyRedirects:
                # Tell the user their URL was bad and try a different one
                assert(0)
            except requests.exceptions.UnboundLocalError:
                # Maybe set up for a retry, or continue in a retry loop
                assert(0)
            except requests.exceptions.RequestException as e:
                # catastrophic error. bail.
                print (e)
                assert(0)
            # end of try

            # got someting from the web, check what it is   
            r.encoding = encoding
            if ('查無資料' in r.text): break # break while loop, move on to next REPORT_ID
            
            try:
                dfs = pd.read_html(r.text)
            except:
                # special wait time
                print (r.text)
                print ("Can't get %s for %d in %d Q%d" %(report_name, co_id, cn_year, quarter))
                print ("Wait %d sec and try again" %(wait_time*10))
                time.sleep(wait_time*10)
            else:
                # standard wait time
                time.sleep(wait_time)
                assert(dfs is not None)
                if len(dfs) >= 5:
                    print('Got %s for %d in %d Q%d' %(report_name, co_id, cn_year, quarter))
                    break #while loop
                print(r.url)
                print(r.text)
                print(co_id, cn_year, quarter)
                if ('下載案例文件' in r.text):
                    #bad luck, saw this on co_id=3454, year=2015, quarter=2
                    if (co_id == 3454 and cn_year == 104 and quarter == 2): return None
                    if (co_id == 4906 and cn_year == 103 and quarter == 2): return None

                    url1 = 'http://mops.twse.com.tw/server-java/FileDownLoad'
                    payload = {
                        'step':'9',
                        'functionName':'t164sb01',
                        'report_id':'C',
                        'co_id':co_id,
                        'year':cn_year+1911,
                        'season':quarter }
                    try:
                        r1 = requests.post(url1, data=payload)
                        if r1.status_code == 200:
                            print(r1.headers)
                            if 'content-disposition' in r1.headers:
                                print("File Attached")
                            else:
                                print("No File Attached")
                            downloaded_filename = r1.headers['content-disposition'].split('"')[1]
                            with open(downloaded_filename, "wb") as f:
                                f.write(r1.content)
                                print('{0} downloaded'.format(downloaded_filename))
                        else:
                            print('Response Status Code: {0}'.format(r1.status_code))
                    except requests.exceptions.ConnectionError:
                        print('File not available')
                    dfs = pd.read_html(r1.text)
                    if (dfs is not None): break
                    assert(0) # what's this?
                # end of fileDownload check
            # end of dfs read
        # end of while(dfs is None)
        if (dfs is not None): return dfs
    # continue to next REPORT_ID
    return dfs



# get tables from the xls or web
def individual_quarterly_report(table_now, co_id, co_name, year, quarter):
    global num_web_access
    # 假如是西元，轉成民國
    if year > 1900: cn_year = year - 1911
    else : cn_year = year

    if (year <= 2012) :
        (url, index_name, encoding) = before_IFRS_url_table[table_now] # before IFRS
        r = requests.post(url, {
                'step':1,
                'firstin':1,
                'off':1,
                'co_id':co_id,
                'year':cn_year,
                'season':'0'+str(quarter)
            })
        r.encoding = encoding
        try:
            dfs = pd.read_html(r.text)
        except:
            print ("[Stock Crawler]: can't get web data from %s" %(url))
            pass # normally, the page will say: 查無資料
        time.sleep(wait_time*((co_id % 10 == 0)*5+1))
    
        if (len(dfs[0].index[(dfs[0][0] == index_name)]) <= 0) :
            print ("[Stock Crawler]: can't find %s" %index_name)
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
            r.encoding = encoding
            try:
                print ("[Stock Crawler]: can't get web data from %s" %(url))
                dfs = pd.read_html(r.text)
            except:
                pass
            time.sleep(wait_time*((co_id % 10 == 0)*5+1))

        # get the tables that are meaningful
        # df = pd.concat([df_now for df_now in dfs if (df_now.shape[1] <= 50 and df_now.shape[0] > 3)])
        df = dfs[0]
        # remove invalid rows
        df = df[df.notnull()]

        if (len(df.index[(df[0] == index_name)]) > 0):
            row_idx_of_columns = df.index[(df[0] == index_name)][0]
            df.columns = df.iloc[row_idx_of_columns]
            # move header row to the row next to '會計科目'
            df = df[(row_idx_of_columns+1) :]
            # change the index to '會計科目'
            df = df.set_index(index_name)
            # drop last row
            df = df.drop(df.index[df.index.size-1])
            # drop empty columns
            df = df.dropna(axis=1, how='all')
        else:
            print('url = %s' %url)
            print('r.text = %s' %r.text)
            print ("[Stock Crawler]: still can't find %s" %index_name)
            assert(0) # why are we here?
    else :
        # after 2012, check if we can read it from previous log
        (url, index_name, encoding) = after_IFRS_url_table[table_now] # after IFRS
        path_name = './'+str(year)+'/Q'+str(quarter)+'/'
        xls_file_name = path_name+'quarterly_IFRS_'+ str(year)+'Q'+str(quarter)+'_'+str(co_id)+'.xlsx'
        xls_file = Path(xls_file_name)
        df = None
        if (xls_file.is_file()) :
            xls_reader = pd.ExcelFile(xls_file_name)
            try:
                df = xls_reader.parse(table_now)
                print("Got table_now for %s from %s without crawling web!!" %(table_now, xls_file_name))
            except:
                pass
        if (df is None):
            # can't get from previous xls, start crawling web
            dfs = CrawForIFRSreport(url, encoding, co_id, cn_year, quarter)
            if (dfs is None): return None
            # get the tables that are meaningful
            assert(len(dfs) >= 5)
            for dfs_idx in range(1,4):
                if (dfs[dfs_idx].iloc[1][0] == table_now): break
            assert(dfs[dfs_idx].iloc[1][0] == table_now)

            # save all tables into xls file
            xls_writer = pd.ExcelWriter(xls_file_name, engine='xlsxwriter')
            for df_idx,df_now in enumerate(dfs):
                # skip the tables that are too short
                if (df_now.shape[0] < 2): continue
                if (df_idx <= 3):
                    sheet_name = df_now.iloc[1][0]
                elif (df_idx == 4):
                    sheet_name = '當期權益變動表'
                elif (df_idx == 5):
                    sheet_name = '去年同期權益變動表'
                else:
                    sheet_name = 'df_%d' %df_idx
                df_now.to_excel(xls_writer, sheet_name=sheet_name)
            xls_writer.save()
                
            df = dfs[dfs_idx]
        # end of web crawling
        
        # remove invalid rows
        df = df[df.notnull()]

        # set columns and index to dataframe
        if (len(df.index[(df[0] == index_name)]) > 0) :
            row_idx_of_columns = df.index[(df[0] == index_name)][0]
            df.columns = df.iloc[row_idx_of_columns]
            # move header row to the row next to '會計項目'
            df = df[(row_idx_of_columns+1) :]
            # change the index to '會計項目'
            df = df.set_index(index_name)
            # drop empty columns
            df.dropna(axis=1, how='all')
        else:
            print(df[0], index_name, [(df[0] == index_name)], df.index, df.index[(df[0] == index_name)])
            assert(0)

    assert(df is not None)
    return df

    