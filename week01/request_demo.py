# -*- coding: utf-8 -*-
# @Author: HoRan.li
# @Date:   2020-06-28 16:10:05
# @Last Modified by:   HoRan.li
# @Last Modified time: 2020-06-28 16:40:23
# @E-mail: laken_phil@163.com

"""
Function Information:
获取猫眼电影
"""

import requests
from bs4 import BeautifulSoup as bs

user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36'
cookie = 'uuid_n_v=v1; uuid=BD68AAB0B91011EA95E3C92BBC8FF01956AB1662EF4740ECA014800E3F96BE2E; _csrf=23f5d181e5cc3f6743b9e7adc551653cd5418e55f77e5be2e38114a56de85551; mojo-uuid=455532745944604cc6ed0cf1fdf2fb32; _lxsdk_cuid=172f9d263bb8c-0fe19d4dc839dc-31617402-1aeaa0-172f9d263bcc8; _lxsdk=BD68AAB0B91011EA95E3C92BBC8FF01956AB1662EF4740ECA014800E3F96BE2E; mojo-session-id={"id":"e52f9514f3c758dc0f0241e2c3b1f378","time":1593334965184}; Hm_lvt_703e94591e87be68cc8da0da7cbd0be2=1593329214,1593336441,1593336455; Hm_lpvt_703e94591e87be68cc8da0da7cbd0be2=1593336455; __mta=216371497.1593329214545.1593336441550.1593336455593.22; _lxsdk_s=172fa2a2337-b8d-a14-6ff%7C%7C19; mojo-trace-id=11'
header = {'user-agent':user_agent, 'Cookie':cookie}

myurl = 'https://maoyan.com/films?showType=3&sortId=3'

response = requests.get(myurl, headers = header)

bs_info = bs(response.text, 'html.parser')
print(bs_info)
# for 解析html
for tags in bs_info.find_all('div', attrs={'class':'movie-hover-info'}):
    print(tags)
    for divTag in tags.find_all('div', attrs={'class':'movie-hover-title'}):
        print(divTag)
        print(divTag.get('title'))

