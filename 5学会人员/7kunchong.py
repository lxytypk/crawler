import re
import pandas as pd
import os
import requests
from lxml import etree
from openpyxl import load_workbook
from selenium import webdriver
from time import sleep

'''中国昆虫学会'''
total_url='http://entsoc.ioz.cas.cn/zzjg/'
university='中国昆虫学会'
result=[]
work_dir=r"C:\Users\Lenovo\Desktop\0314"

headers={
'User-Agent': #'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:134.0) Gecko/20100101 Firefox/134.0'
}

'''各学院url'''
#实例化一个浏览器对象
bro=webdriver.Firefox()
#让浏览器发起一个指定url对应的请求
bro.get('http://entsoc.ioz.cas.cn/zzjg/')
#获取浏览器当前页面的页面源码数据
page_text=bro.page_source
print(page_text)

# response = requests.get(url=total_url, headers=headers)
# response.encoding = 'utf-8'
tree = etree.HTML(page_text)


div_list=tree.xpath('//table[@class="Table"]/tbody/tr[position()>1]')

for d in div_list:
    sleep(1)
    name=d.xpath('./td[2]/p/text()')[0].strip()
    job = d.xpath('./td[3]/p/text()')[0]

    institution= d.xpath('./td[8]/p/text()')[0]

    dict = {
        '学会理事URL': total_url,
        '学会名称': university,
        '姓名': name,
        '职位': job,
        '机构': institution,
        '邮箱': '',
        '任职开始年份': '2022年11月21日',
        '任职结束年份': ''
    }
    result.append(dict)
    print('--------------------------------------------')

'''续写Excel文件'''
ff=pd.DataFrame(result)
file_path = f"{work_dir}/lxy0314.xlsx"
d1=pd.read_excel(file_path)
d1 = pd.concat([d1, ff], ignore_index=True)	# 合并数据

book = load_workbook(file_path)  # 该文件必须存在,并且该语句必须在 with pd.ExcelWriter() 之前
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    writer.book = book
    d1.to_excel(writer, sheet_name="Sheet1",index=False)  # 重写sheet