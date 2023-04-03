import requests
from bs4 import BeautifulSoup

def get_name(in_stock_num):
    url = "https://goodinfo.tw/tw/StockDetail.asp?STOCK_ID="
    url += str(in_stock_num)
    
    # avoid protection of goodinfo
    headers = {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36"
    }
    res = requests.get(url, headers=headers)

    # to Chinese
    res.encoding = "utf-8"

    soup = BeautifulSoup(res.text, "lxml")
    link = soup.find("tr", class_="bg_h0 fw_normal")
    stock_num_name = link.select_one("a").getText()
    stock_name = stock_num_name[5:]

    return stock_name

def get_price(in_stock_num):
    url = "https://tw.stock.yahoo.com/quote/"
    url += str(in_stock_num)
    
    # avoid protection of goodinfo
    res = requests.get(url, timeout=5)

    soup = BeautifulSoup(res.text, "lxml").select("#atomic .Fz\(32px\)")
    for price in soup:
        price = price.text

    return price

if __name__ == "__main__":
    #get_name(2851)
    get_price(2851)