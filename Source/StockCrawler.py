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
ref_xls_file_name = "HGcomplete.xls"
ref_xls_file = Path(ref_xls_file_name)
if not ref_xls_file.is_file() :
    exit("Please prepare %s for target stock list" %ref_xls_file_name)
# xls file exists, read tables from xls
ref_xls_reader = pd.ExcelFile(ref_xls_file_name, encoding='utf-8')
ref_df = ref_xls_reader.parse("Dashboard")
ref_xls_reader.close()
ref_df = ref_df.set_index('代號')
target_stocks = list(ref_df.index)
#target_stocks = [1101]

# get 月成交資訊 trading record for interested stocks
month_range = {
#    2013: range(1,13),
#    2014: range(1,13),
#    2015: range(1,13),
#    2016: range(1,13),
#    2017: range(1,13),
    2018: range(1,2) # as of 2018/03/11 no data for Q1 yet
}
update_latest_monthly = False # flag to update un-finished or latest monthly reports
raw_individual_monthly_trading_table = {}
year_range = sorted(list(month_range.keys()))
for year in year_range :
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
        sheet_name = "%s_%d" %(month_string, co_id)
        xls_file_name = path_name + month_string + '_' + str(co_id) + '.xlsx'
        xls_file_now = Path(xls_file_name)
        if (xls_file_now.is_file()) :
            # xls file exists, read tables from xls
            xls_reader = pd.ExcelFile(xls_file_name)
            try:
                df_now = xls_reader.parse(sheet_name)
            except:
                sheet_name_with_co_name = "%s_%d_%s" %(month_string, co_id, co_name)
                try:
                    df_now = xls_reader.parse(sheet_name_with_co_name)
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


quarter_range = {
        2013: range(1,5),
#        2014: range(2,5),
#        2015: range(2,5),
#        2016: range(1,5),
#        2017: range(1,5), # as of 2018/03/18 no data for Q4'17 yet
#        2018: range(1,1) # as of 2018/03/11 no data for Q1 yet
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

iq_tables = {}

iq_dict = {
    '資產負債表': 'asset_',
    '現金流量表': 'cashflow_',
    '綜合損益表': 'pnl_'
}

for iqd_now in iq_dict.keys() :
    xls_suffix = iq_dict[iqd_now]
    for year in quarter_range.keys() :
        for quarter in quarter_range[year] :
            path_name = './'+str(year)+'/Q'+str(quarter)+'/'
            directory = os.path.dirname(path_name)
            try:
                os.stat(directory)
            except:
                os.mkdir(directory)      
            for co_id in target_stocks:
                co_name = ref_df.loc[co_id, '公司']
                if (co_id in IPO_Date.keys()) : # this stock is still young
                    if (year < IPO_Date[co_id][0] or (year == IPO_Date[co_id][0] and quarter <= ((IPO_Date[co_id][1]/3)+1))) :
                        print("%s is not IPO yet in %sQ%s" %(co_name, year, quarter))
                        continue;
                table_name = 'quarterly_individual_'+xls_suffix+str(year)+str(quarter)+'_'+str(co_id)
                xls_file_name = path_name + table_name + '.xlsx'
                sheet_name = xls_suffix+str(year)+str(quarter)+'_'+str(co_id)
                xls_file = Path(xls_file_name)
                if (xls_file.is_file()) :
                    # xls file exists, read tables from xls
                    print("Reading %s %d %s from %s@%s" %(iqd_now, co_id, co_name, sheet_name, xls_file_name))
                    xls_reader = pd.ExcelFile(xls_file_name)
                    df_now = xls_reader.parse(sheet_name)
                    # use next line as column if it wasn't properly set
                    if (df_now.columns[0] == 0): df_now.columns = df_now.iloc[0]
                    if ('會計項目' in df_now.columns):
                        df_now = df_now.set_index('會計項目')
                    else :
                        assert('會計科目' in df_now.columns)
                        df_now = df_now.set_index('會計科目')
                    xls_reader.close()
                    df_now = df_now.drop_duplicates(keep='last')
                    iq_tables[table_name] = df_now
                else :
                    print("Crawling %s for %d %s %dQ%d" %(iqd_now, co_id, co_name, year, quarter))
                    df_now = individual_quarterly_report(iqd_now, co_id, co_name, year, quarter)
                    if (df_now is not None) :
                        print("Success, writing %s to %s" %(sheet_name, xls_file_name))
                        xls_writer = pd.ExcelWriter(xls_file_name, engine='xlsxwriter')
                        df_now.to_excel(xls_writer, sheet_name=sheet_name)
                        xls_writer.save()
                        df_now = df_now.drop_duplicates(keep='last')
                        iq_tables[table_name] = df_now
                    else :
                        print("Error: can't get %s for %s" %(sheet_name, xls_file_name))
            # done looping companies
        # done looping quarter
    # done looping year
# done looping tables



print ("Finished DataFrame preparation! Next step is to prepare data into Excel Stock model!")

summary_columns = {
        "現金及約當現金總額": ('asset_', "現金及約當現金"),
        "現金及約當現金合計": ('asset_', "現金及約當現金"),
        "應收票據淨額": ("asset_", "應收"),
        "應收帳款淨額": ("asset_", "應收"),
        "存貨合計":("asset_", "存貨"),
        "無形資產合計": ("asset_", "無形資產"),
        "營業活動之淨現金流入（流出）":("cashflow_", "淨現金流"),
        "投資活動之淨現金流入（流出）":("cashflow_", "淨現金流"),
        "籌資活動之淨現金流入（流出）":("cashflow_", "淨現金流")}

# prepare columns of summary table
targ_col_list = [summary_columns[idx_now][1] for idx_now in summary_columns.keys()]
targ_col_list += summary_columns.keys()
targ_col_set = set(targ_col_list)
targ_idx_set = [str(co_id) for co_id in target_stocks]
for year in quarter_range.keys() :
    for quarter in quarter_range[year] :
        df_new = pd.DataFrame(0, index=targ_idx_set, columns=['公司']+list(targ_col_set))
        for co_id in target_stocks:
            target_idx = [str(co_id)]
            df_new.loc[target_idx, '公司'] = ref_df.loc[co_id, '公司']
            # get data from source tables
            for src_idx in summary_columns.keys():
                (suffix_now, target_col) = summary_columns[src_idx]
                table_name = 'quarterly_individual_'+suffix_now+str(year)+str(quarter)+'_'+str(co_id)
                print("Working on %s in %s" %(src_idx, table_name))
                if (table_name not in iq_tables.keys()):
                    if (co_id in IPO_Date.keys()): # this stock is still young
                        if (year*12+quarter*3) <= (IPO_Date[co_id][0]*12+IPO_Date[co_id][1]*3+3): continue
                    assert(0) # why don't we have this table?
                df_now = iq_tables[table_name]
                if (src_idx in df_now.index):
                    value_now = float(df_now[df_now.columns[0]][src_idx])
                    # cashflow numbers are quarterly cumulated
                    if (suffix_now == 'cashflow_' and quarter > 1):
                        table_last = 'quarterly_individual_'+suffix_now+str(year)+str(quarter-1)+'_'+str(co_id)
                        if (table_last in iq_tables.keys()):
                            df_last = iq_tables[table_last]
                            if (src_idx in df_last.index):
                                value_now -= float(df_last[df_last.columns[0]][src_idx])
                    df_new.loc[target_idx, target_col] += value_now
                    if (target_col != src_idx):
                        df_new.loc[target_idx, src_idx] = value_now
        csv_file_name = 'cashflow_'+str(year)+str(quarter)+'.csv'
        df_new.to_csv(csv_file_name, encoding='utf_8_sig')

    