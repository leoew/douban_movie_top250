# -*- coding: utf-8 -*-
import requests
import re
import xlwt

base_url = 'https://movie.douban.com/top250'


def get_url(page):
    if page == 0:
        url = base_url
    else:
        url = 'https://movie.douban.com/top250?start=' + str(25*page) + '&filter='
    return url


def get_html(url):
    r = requests.get(url)
    return r.content


def get_info(content):
    pattern = re.compile('<span class="title">(.*?)</span>.*?'
                         '<span class=.*?>&nbsp;/&nbsp;(.*?)</span>.*?'
                         '(\d\d\d\d).*?'
                         '<span class="rating_num" property="v:average">(.*?)</span>.*?'
                         '<span>(.*?)人评价</span>.*?', re.S)
    info = re.findall(pattern, content)
    return info


def write(infos, i):
    for info in infos:
        sheet1.write(i, 0, info[0].decode('utf-8'))
        sheet1.write(i, 1, info[1].decode('utf-8').replace('&#39;', "'"))
        sheet1.write(i, 2, info[2].decode('utf-8'))
        sheet1.write(i, 3, info[3].decode('utf-8'))
        sheet1.write(i, 4, info[4].decode('utf-8'))

        i += 1

f = xlwt.Workbook()
sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
for page_num in range(10):
    page_url = get_url(page_num)
    html = get_html(page_url)
    items = get_info(html)

    write(items, page_num*25)
f.save('top250.xls')

