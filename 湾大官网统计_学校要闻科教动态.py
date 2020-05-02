# 北部湾大学官网数据统计
# 支持学校要闻和科教动态栏目
# 默认栏目为学校要闻，需要设置起始目录页和终止目录页
# 自动导出到Excel

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import colorama
import random
import re

base_url = "http://www.bbgu.edu.cn"

news = "xxyw"    # 学校要闻1038,科技动态1040
if news == "xxyw":
    news_type = "1038"
    news_format = "xxyw"
elif news == "kjdt":
    news_type = "1040"
    news_format = "kjdt"


# 设置起始页码和中止页码
start_pages = 320
end_pages = 341

# 新建Excel
wbk = Workbook()

headers = {'Accept': 'text/html, application/xhtml+xml, image/jxr, */*',
           'Accept - Encoding': 'gzip, deflate',
           'Accept-Language': 'zh-Hans-CN, zh-Hans; q=0.5',
           'Connection': 'Keep-Alive',
           "Referer": "http://www.qzhu.edu.cn/",
           "Upgrade-Insecure-Requests": "1",
           "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36"}


def get_views(clickid, owner, clicktype):
    # 获取阅读量
    views = requests.get(base_url+"/system/resource/code/news/click/dynclicks.jsp?clickid=" +
                         clickid+"&owner=1364673649&clicktype="+clicktype+"", headers=headers).text
    return views


def get_content(news_url):
    # 获取网页内容并返回 bs对象
    url = news_url
    request = requests.get(url, headers=headers)
    request.encoding = "utf-8"
    content = request.content
    return BeautifulSoup(content, "lxml")


def get_news_pages():
    # 获取每个新闻页目录，返回列表
    urls = []
    for i in range(start_pages, end_pages):
        urls.append(base_url+"/xxxw/"+news_format+"/"+str(i)+".htm")
    urls.append(base_url+"/xxxw/"+news_format+".htm")
    return urls


row = ["标题", "来源", "日期", "阅读量"]
sheet = wbk.create_sheet(news_format, index=1)
sheet.append(row)
# 创建Excel表并添加表头

for page in get_news_pages():
    urls = BeautifulSoup(requests.get(page).content,
                         "lxml").find_all("a", class_="block oh c666 fs14")
    # 获取每个目录页的新闻页网址
    for i in urls:
        if "../.." in i["href"]:  # 起始页网址为../ 其他为../..，统一格式化网址
            news_url = i["href"].replace("../..", base_url)
        else:
            news_url = i["href"].replace("..", base_url)

        news_id = news_url.split("/")[-1].split(".")[0]  # 获取新闻ID
        soup = get_content(news_url)

        title = soup.find("div", class_="c_title pt20 mt5").h1.text
        branch = soup.find("span", class_="branch").text.split("：")[1].strip()
        date = soup.find("span", class_="time").text.split("：")[1]
        views = get_views(news_id, "1364673649", "wbnews")

        color_code = random.randint(1, 7)
        colors = ["BLACK", "RED", "GREEN", "YELLOW",
                  "BLUE", "MAGENTA", "CYAN", "WHITE"]
        # 获取随机颜色方便输出

        sheet.append([title, branch, date, views])
        wbk.save("官网统计-"+news_format+"xlsx")
        # 添加每行并保存Excel

        print(eval("colorama.Fore."+colors[color_code])+title+"\n来源:" +
              branch+"---日期："+date+"---阅读量："+views)
        print(colorama.Fore.RESET)
        # 输出并重置颜色
