# -*- coding: utf-8 -*-
"""
Spyder Editor
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import time
from pathlib import Path
import os

from combined_quarterly_statement import combined_quarterly_statement
from combined_monthly_revenue_report import combined_monthly_revenue_report
from individual_quarterly_report import individual_quarterly_report
from individual_monthly_trading_record import individual_monthly_trading_record

# main configurations
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

# read target stock list
ref_xls_file_name = "HG交易系統.xlsx"
ref_xls_file = Path(ref_xls_file_name)
if (not ref_xls_file.is_file()) :
    exit("Please prepare %s for target stock list" %ref_xls_file_name)
# xls file exists, read tables from xls
ref_xls_reader = pd.ExcelFile(ref_xls_file_name)
ref_df = ref_xls_reader.parse("儀表板")
ref_xls_reader.close()
ref_df = ref_df.set_index("代號")
target_stocks = list(ref_df.index)

quarter_range = {
        2013: range(1,5),
        2014: range(1,5),
        2015: range(1,5),
        2016: range(1,5),
        2017: range(1,4),
        2018: range(1,1) # as of 2018/03/11 no data for Q1 yet
}


combined_quarterly_tables = [
        'SII綜合損益彙總表',
        'OTC綜合損益彙總表',
        'SII資產負債彙總表',
        'OTC資產負債彙總表',
        'SII營益分析彙總表',
        'OTC營益分析彙總表']

combined_monthly_tables = [
        'SII每月營業收入彙總表',
        'OTC每月營業收入彙總表']

raw_quarterly_table = {}
raw_monthly_table = {}
for year in quarter_range.keys() :
    for quarter in quarter_range[year] :
        path_name = './'+str(year)+'/Q'+str(quarter)+'/'
        directory = os.path.dirname(path_name)
        try:
            os.stat(directory)
        except:
            os.mkdir(directory)
        xls_file_name = path_name+'financial_statement_'+str(year)+'Q'+str(quarter)+'.xlsx'
        xls_file = Path(xls_file_name)
        if (xls_file.is_file()) :
            # xls file exists, read tables from xls
            xls_reader = pd.ExcelFile(xls_file_name)
            for table_now in combined_quarterly_tables:
                sheet_name = table_now + '_' + str(year)+ 'Q'+str(quarter)
                print("Reading %s from %s" %(sheet_name, xls_file_name))
                try:
                    df_now = xls_reader.parse(sheet_name)
                    df_now = df_now.set_index('公司代號')
                except:
                    df_now = None
                raw_quarterly_table[sheet_name] = df_now
            for month in range((quarter-1)*3+1, (quarter-1)*3+4) :
                for table_now in combined_monthly_tables:
                    sheet_name = table_now + '_' + str(year) + '_' + str(month)
                    print("Reading %s from %s" %(sheet_name, xls_file_name))
                    try:
                        df_now = xls_reader.parse(sheet_name)
                        df_now = df_now.set_index('公司代號')
                    except:
                        df_now = None
                    raw_monthly_table[sheet_name] = df_now
            xls_reader.close()
        else :
            # read tables from web and save to xls
            xls_writer = pd.ExcelWriter(xls_file_name, engine='xlsxwriter')
            for table_now in combined_quarterly_tables:
                sheet_name = table_now + '_' + str(year)+ 'Q'+str(quarter)
                df_now = combined_quarterly_statement(year, quarter, table_now)
                raw_quarterly_table[sheet_name] = df_now
                if (df_now is None): continue
                print("Writing %s to %s" %(sheet_name, xls_file_name))
                df_now.to_excel(xls_writer, sheet_name=sheet_name)
            for month in range((quarter-1)*3+1, (quarter-1)*3+4) :
                for table_now in combined_monthly_tables:
                    sheet_name = table_now + '_' + str(year) + '_' + str(month)
                    df_now = combined_monthly_revenue_report(year, month, table_now)
                    raw_monthly_table[sheet_name] = df_now
                    if (df_now is None): continue
                    print("Writing %s to %s" %(sheet_name, xls_file_name))
                    df_now.to_excel(xls_writer, sheet_name=sheet_name)
            xls_writer.save()
    # end of quarter loop
# end of year loop


# Get quarterly 資產負債表
raw_individual_quarterly_asset = {}
#for co_id in list(raw_monthly_table[sheet_name]['公司代號'])
for year in quarter_range.keys() :
    raw_individual_quarterly_asset[str(year)] = [[], [], [], []]
    for quarter in quarter_range[year] :
        path_name = './'+str(year)+'/Q'+str(quarter)+'/'
        directory = os.path.dirname(path_name)
        try:
            os.stat(directory)
        except:
            os.mkdir(directory)      
        for co_idx, co_id in enumerate(target_stocks) :
            co_name = ref_df.loc[co_id, '公司']
            if (co_id in IPO_Date.keys()) : # this stock is still young
                if (year < IPO_Date[co_id][0] or (year == IPO_Date[co_id][0] and quarter <= ((IPO_Date[co_id][1]/3)+1))) :
                    print("%s is not IPO yet in %sQ%s" %(co_name, year, quarter))
                    continue;
            xls_file_name = path_name+'quarterly_indivisual_asset_'+ str(year)+'Q'+str(quarter)+'_'+str(co_id)+'.xlsx'
            sheet_name = str(co_id) + '_' + co_name
            xls_file = Path(xls_file_name)
            if (xls_file.is_file()) :
                # xls file exists, read tables from xls
                print("Reading %s from %s" %(sheet_name, xls_file_name))
                xls_reader = pd.ExcelFile(xls_file_name)
                df_now = xls_reader.parse(sheet_name)
                # use next line as column if it wasn't properly set
                if (df_now.columns[0] == 0): df_now.columns = df_now.iloc[0]
                if ('會計項目' in df_now.columns):
                    df_now.set_index('會計項目')
                else :
                    assert('會計科目' in df_now.columns)
                    df_now.set_index('會計科目')
                xls_reader.close()
                raw_individual_quarterly_asset[str(year)][quarter-1] = df_now
            else :
                print("Crawling %s for %d:%d %s %dQ%d" %('資產負債表', co_idx, co_id, co_name, year, quarter))
                df_now = individual_quarterly_report('資產負債表', co_id, co_name, year, quarter)
                if (df_now is not None) :
                    print("Success, writing %s to %s" %(sheet_name, xls_file_name))
                    xls_writer = pd.ExcelWriter(xls_file_name, engine='xlsxwriter')
                    df_now.to_excel(xls_writer, sheet_name=sheet_name)
                    xls_writer.save()
                    raw_individual_quarterly_asset[str(year)][quarter-1] = df_now
                else :
                    print("Error: can't get %s for %s" %(sheet_name, xls_file_name))
        # done looping companies
    # done looping quarter
# done looping year


# Get quarterly 現金流量表
raw_individual_quarterly_cashflow = {}
num_crawler_called = 0
#for co_id in list(raw_monthly_table[sheet_name]['公司代號'])
for year in quarter_range.keys() :
    raw_individual_quarterly_cashflow[str(year)] = [[], [], [], []]
    for quarter in quarter_range[year] :
        path_name = './'+str(year)+'/Q'+str(quarter)+'/'
        directory = os.path.dirname(path_name)
        try:
            os.stat(directory)
        except:
            os.mkdir(directory)      
        for co_idx, co_id in enumerate(target_stocks) :
            co_name = ref_df.loc[co_id, '公司']
            if (co_id in IPO_Date.keys()) : # this stock is still young
                if (year < IPO_Date[co_id][0] or (year == IPO_Date[co_id][0] and quarter <= ((IPO_Date[co_id][1]/3)+1))) :
                    print("%s is not IPO yet in %sQ%s" %(co_name, year, quarter))
                    continue;
            xls_file_name = path_name+'quarterly_indivisual_cashflow_'+ str(year)+'Q'+str(quarter)+'_'+str(co_id)+'.xlsx'
            sheet_name = str(co_id) + '_' + co_name
            xls_file = Path(xls_file_name)
            if (xls_file.is_file()) :
                # xls file exists, read tables from xls
                print("Reading %s from %s" %(sheet_name, xls_file_name))
                xls_reader = pd.ExcelFile(xls_file_name)
                df_now = xls_reader.parse(sheet_name)
                if ('會計項目' in df_now.columns):
                    df_now.set_index('會計項目')
                else :
                    assert('會計科目' in df_now.columns)
                    df_now.set_index('會計科目')
                xls_reader.close()
                raw_individual_quarterly_cashflow[str(year)][quarter-1] = df_now
            else :
                print("Crawling %s for %d:%d %s %dQ%d" %('現金流量表 ', co_idx, co_id, co_name, year, quarter))
                df_now = individual_quarterly_report('現金流量表', co_id, co_name, year, quarter)
                if (df_now is not None) :
                    print("Success, writing %s to %s" %(sheet_name, xls_file_name))
                    xls_writer = pd.ExcelWriter(xls_file_name, engine='xlsxwriter')
                    df_now.to_excel(xls_writer, sheet_name=sheet_name)
                    xls_writer.save()
                    raw_individual_quarterly_cashflow[str(year)][quarter-1] = df_now
                else :
                    print("Error: can't get %s for %s" %(sheet_name, xls_file_name))
                num_crawler_called += 1
            if (num_crawler_called % 30 == 29):
                print("Pause a while to avoid rejection")
                time.sleep(15)
        # done looping companies
    # done looping quarter
# done looping year

"""
# get 月成交資訊 trading record for interested stocks
month_range = {
#        2013: range(1,13),
#        2014: range(1,13),
#        2015: range(1,13),
#        2016: range(1,13),
        2017: range(1,13),
        2018: range(1,1) # as of 2018/03/11 no data for Q1 yet
}
update_latest_monthly = False # flag to update un-finished or latest monthly reports
raw_individual_monthly_trading_table = {}
for year in month_range.keys() :
    path_name = './'+str(year)+'/'
    directory = os.path.dirname(path_name)
    try:
        os.stat(directory)
    except:
        os.mkdir(directory)
    month_string = "trading_%4d0101" %year
    for co_id in target_stocks:
        if (co_id in IPO_Date.keys()): # this stock is still young
            if (year < IPO_Date[co_id][0]):
                continue
        co_name = ref_df.loc[co_id, '公司']
        sheet_name = "%s_%d_%s" %(month_string, co_id, co_name)
        xls_file_name = path_name + month_string + '_' + str(co_id) + '.xlsx'
        xls_file_now = Path(xls_file_name)
        if (xls_file_now.is_file()) :
            # xls file exists, read tables from xls
            xls_reader = pd.ExcelFile(xls_file_name)
            try:
                df_now = xls_reader.parse(sheet_name)
            except:
                df_now = None
            # end of co_id loop
            xls_reader.close()
            if (df_now is not None):
                if ('月份' in df_now.columns):
                    df_now = df_now.set_index('月份')
                else:
                    assert('月' in df_now.columns)
                    if (df_now.iloc[0][0] == '月'):
                        df_now = df_now.drop(index=0)
                        xls_writer = pd.ExcelWriter(xls_file_name, engine='xlsxwriter')
                        df_now.to_excel(xls_writer, sheet_name=sheet_name)
                        xls_writer.save()
                    df_now = df_now.set_index('月')
                if (df_now.index[-1] == 12 or df_now.index[-1] == str(12) or \
                    (year == year_range[-1] and update_latest_monthly == False)) :
                    print("Got %s from xls" %sheet_name)
                    # save df_now to the array
                    raw_individual_monthly_trading_table[sheet_name] = df_now
                    continue # next co_id
        # no previous data, read tables from web and save to xls
        df_now = individual_monthly_trading_record(co_id, co_name, year, ref_df)
        raw_individual_monthly_trading_table[sheet_name] = df_now
        if (df_now is not None):
            print("Writing %s to %s" %(sheet_name, xls_file_name))
            xls_writer = pd.ExcelWriter(xls_file_name, engine='xlsxwriter')
            df_now.to_excel(xls_writer, sheet_name=sheet_name)
            xls_writer.save()
        else:
            print("Failed getting %s for %s" %(sheet_name, xls_file_name))
    # for co_id
# end of year loop
"""

print ("Finished DataFrame preparation! Next step is to prepare data into Excel Stock model!")

