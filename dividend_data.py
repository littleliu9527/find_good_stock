from math import fabs
import pandas as pd
import requests
import openpyxl
import datetime
from bs4 import BeautifulSoup

save_version = "0.0.1"
sheet0_name = "version"
sheet1_name = "following dividend stock"
sheet2_name = "dividend data"
sheet3_name = "possibility of dividend data"
sheet4_name = "history dividend date"

# avoid protection of goodinfo
goodinfo_headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36"
}

def create_excel():
    wb = openpyxl.Workbook()

    # add version
    sheet = wb.worksheets[0]
    sheet.title = sheet0_name
    sheet['A1'] = "version"
    sheet['B1'] = save_version

    # create file name with current time
    current_time = datetime.datetime.now()
    file_name = "stock_test"
    file_name += "_"
    file_name += str(current_time.year)
    file_name += "_"
    file_name += str(current_time.month).zfill(2)
    file_name += "_"
    file_name += str(current_time.day).zfill(2)
    file_name += "_"
    file_name += str(current_time.hour).zfill(2)
    file_name += "_"
    file_name += str(current_time.minute).zfill(2)
    file_name += "_"
    file_name += str(current_time.second).zfill(2)
    file_name += ".xlsx"

    # save file
    wb.save(file_name)

    return file_name

def get_following_dividend_data():
    url = "https://goodinfo.tw/tw/StockDividendScheduleList.asp?MARKET_CAT=%E5%85%A8%E9%83%A8&"+ \
        "YEAR=%E5%8D%B3%E5%B0%87%E9%99%A4%E6%AC%8A%E6%81%AF&INDUSTRY_CAT=%E5%85%A8%E9%83%A8"
    
    res = requests.get(url, headers=goodinfo_headers, timeout=5)

    # to Chinese
    res.encoding = "utf-8"

    soup = BeautifulSoup(res.text,"lxml")
    data = soup.select_one("#tblDetail")
    dfs = pd.read_html(str(data))
    df = dfs[0]

    return df

def get_dividend_data():
    url = "https://goodinfo.tw/tw2/StockList.asp?MARKET_CAT=%E8%87%AA%E8%A8%82%E7%AF%A9%E9%81%B8&"+ \
        "INDUSTRY_CAT=%E6%88%91%E7%9A%84%E6%A2%9D%E4%BB%B6&"+ \
        "FL_ITEM0=%E9%80%A3%E7%BA%8C%E9%85%8D%E7%99%BC%E7%8F%BE%E9%87%91%E8%82%A1%E5%88%A9%E6%AC%A1%E6%95%B8&"+ \
        "FL_VAL_S0=8&FL_VAL_E0=&FL_ITEM1=&FL_VAL_S1=&FL_VAL_E1=&FL_ITEM2=&FL_VAL_S2=&FL_VAL_E2=&"+ \
        "FL_ITEM3=&FL_VAL_S3=&FL_VAL_E3=&FL_ITEM4=&FL_VAL_S4=&FL_VAL_E4=&FL_ITEM5=&FL_VAL_S5=&"+ \
        "FL_VAL_E5=&FL_ITEM6=&FL_VAL_S6=&FL_VAL_E6=&FL_ITEM7=&FL_VAL_S7=&FL_VAL_E7=&FL_ITEM8=&"+ \
        "FL_VAL_S8=&FL_VAL_E8=&FL_ITEM9=&FL_VAL_S9=&FL_VAL_E9=&FL_ITEM10=&FL_VAL_S10=&FL_VAL_E10=&"+ \
        "FL_ITEM11=&FL_VAL_S11=&FL_VAL_E11=&FL_RULE0=%E7%94%A2%E6%A5%AD%E9%A1%9E%E5%88%A5%7C%7CETF&"+ \
        "FL_RULE_CHK0=T&FL_RULE1=&FL_RULE2=&FL_RULE3=&FL_RULE4=&FL_RULE5=&"+ \
        "FL_RANK0=%E8%82%A1%E5%88%A9%E6%8E%92%E8%A1%8C%7C%7C%E9%99%A4%E6%81%AF%E4%BA%A4%E6%98%93%E6%97%A5%7C%7C500&"+ \
        "FL_RANK1=&FL_RANK2=&FL_RANK3=&FL_RANK4=&FL_RANK5=&FL_FD0=&FL_FD1=&FL_FD2=&FL_FD3=&FL_FD4=&"+ \
        "FL_FD5=&FL_SHEET=%E8%82%A1%E5%88%A9%E6%94%BF%E7%AD%96%E7%99%BC%E6%94%BE%E5%B9%B4%E5%BA%A6_%E8%BF%91N%E5%B9%B4%E8%82%A1%E5%88%A9%E4%B8%80%E8%A6%BD&"+ \
        "FL_SHEET2=%E7%8F%BE%E9%87%91%2B%E8%82%A1%E7%A5%A8%E8%82%A1%E5%88%A9&"+ \
        "FL_MARKET=%E4%B8%8A%E5%B8%82%2F%E4%B8%8A%E6%AB%83&FL_QRY=%E6%9F%A5++%E8%A9%A2"
    
    res = requests.get(url, headers=goodinfo_headers, timeout=5)

    # to Chinese
    res.encoding = "utf-8"

    soup = BeautifulSoup(res.text,"lxml")
    data = soup.select_one("#txtStockListData")
    dfs = pd.read_html(str(data))
    df = dfs[1]

    return df

def get_stock_num_from_following_dividend():
    # get stock num column
    df = pd.read_excel(file_name,
                        sheet_name=sheet3_name,
                        usecols="A")
    
    stock_num = df.values

    print(stock_num)

    # find history date
    # select dividend date column
    # transpose
    # append to corresponding stock_num of sheet 4

    return 123

def get_stock_dividend_history_date(stock_num):
    url = "https://goodinfo.tw/tw/StockDividendSchedule.asp?STOCK_ID="
    url += str(stock_num)

    res = requests.get(url, headers=goodinfo_headers, timeout=5)

    # to Chinese
    res.encoding = "utf-8"

def copy_dividend_data():
    # use columns in need
    df = pd.read_excel(file_name,
                        sheet_name=sheet1_name,
                        usecols="C, D, F")
    
    # use rows in need
    start_idx = 3
    df = df.iloc[start_idx:]
    for row_index in range(start_idx, len(df.index)):
         if (row_index % 21 == 0
            or row_index % 21 == 1
            or row_index % 21 == 2):
             df.drop([row_index], axis=0, inplace=True)

    # debug
    #print(df)

    return df

def dividend_data_to_excel(file_name):
    # change to selected folder (if need)
    path = file_name
    writer = pd.ExcelWriter(path, mode='a', engine='openpyxl')

    # sheet 1
    get_following_dividend_data().to_excel(writer, sheet_name=sheet1_name)
    writer.save()   # must save, or cannot use read later
    
    # sheet 2
    get_dividend_data().to_excel(writer, sheet_name=sheet2_name)

    # sheet 3
    copy_dividend_data().to_excel(writer, sheet_name=sheet3_name, index=False)
    writer.save()

    # sheet 4
    get_stock_num_from_following_dividend()

    writer.save()
    writer.close()

if __name__ == "__main__":
    # start time
    start_time = datetime.datetime.now()
    print("engine start")
    print("start time:", start_time)

    # create sheets of excel file
    file_name = create_excel()

    # get dividend data
    dividend_data_to_excel(file_name)

    # end time
    end_time = datetime.datetime.now()
    print("engine complete")
    print("end time:", end_time)
