# https://www.cctanfoundation.org/plus/list.php?tid=35
import re
import pandas as pd
import requests
from lxml import etree

'''谈家桢生命科学奖获得者'''

'''
2024
https://www.cctanfoundation.org/plus/list.php?tid=60&TotalResult=18&PageNo=1
https://www.cctanfoundation.org/plus/list.php?tid=60&TotalResult=18&PageNo=2

2023
https://www.cctanfoundation.org/plus/list.php?tid=59&TotalResult=18&PageNo=1
https://www.cctanfoundation.org/plus/list.php?tid=59&TotalResult=18&PageNo=2
'''
url='https://www.cctanfoundation.org/plus/list.php?tid=35'
headers={
'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0'
}

result=[]
work_dir=r"C:\Users\Lenovo\Desktop\lxy\2award"

response=requests.get(url=url,headers=headers)
response.encoding='utf-8'
tree=etree.HTML(response.text)
# print(response.text)

div_list=tree.xpath('//div[@class="list-box clearfix"]/div')
for div in div_list:
    name=div.xpath('./a/h5/text()')[0]
    try:
        profession = div.xpath('./a/h5/span/text()')[0]
        if not profession.strip():  # 检查 profession 是否为空
            profession = ''  # 设置默认值
    except IndexError:
        profession = ''  # 设置默认值
    reward=div.xpath('./a/p/text()')[0]
    # print(reward)
    # print(profession)
    print(name)
    print('----------------------')
