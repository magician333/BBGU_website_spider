# 北部湾大学官网数据统计
# 媒体聚焦自动统计
# 需要设置起始页和终止目录页

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import colorama
import random
import re

base_url = "http://www.bbgu.edu.cn"


start_pages = 94
end_pages = 99

wbk = Workbook()

headers = {'Accept': 'text/html, application/xhtml+xml, image/jxr, */*',
           'Accept - Encoding': 'gzip, deflate',
           'Accept-Language': 'zh-Hans-CN, zh-Hans; q=0.5',
           'Connection': 'Keep-Alive',
           "Referer": "http://www.qzhu.edu.cn/",
           "Upgrade-Insecure-Requests": "1",
           "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36"}

row = ["标题", "来源", "日期"]
sheet = wbk.create_sheet("媒体聚焦", index=1)
sheet.append(row)


def get_news_pages():
    urls = []
    for i in range(start_pages, end_pages):
        urls.append(base_url+"/xxxw/mtjj/"+str(i)+".htm")
    urls.append(base_url+"/xxxw/mtjj.htm")
    return urls


for page in get_news_pages():
    urls = BeautifulSoup(requests.get(page, headers=headers).content,
                         "lxml").find_all("a", class_="block oh c666 fs14")
    date = BeautifulSoup(requests.get(page, headers=headers).content, "lxml").find_all(
        "span", class_="fr ci fs14")
    for x in zip(urls, date):
        date = x[1].text
        title = x[0]["title"]
        branch = re.search("【.*】", x[0]["title"]).group()[1:-1]
        color_code = random.randint(1, 7)
        colors = ["BLACK", "RED", "GREEN", "YELLOW",
                  "BLUE", "MAGENTA", "CYAN", "WHITE"]
        sheet.append([title, branch, date])
        wbk.save("官网统计-媒体聚焦.xlsx")
        print(eval("colorama.Fore."+colors[color_code])+title+"---来源:" +
              branch+"---日期："+date)
        print(colorama.Fore.RESET)
