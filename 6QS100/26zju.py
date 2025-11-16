import re
import pandas as pd
import requests
from lxml import etree
import os
from openpyxl import load_workbook
from selenium import webdriver
from time import sleep


#实例化一个浏览器对象
bro=webdriver.Firefox()
#让浏览器发起一个指定url对应的请求
bro.get('https://www.zju.edu.cn/xywxw/list.htm')
#获取浏览器当前页面的页面源码数据
page_text=bro.page_source

'''Zhejiang University'''
total_url='https://www.zju.edu.cn/xywxw/list.htm'
university='Zhejiang University'
result=[]
work_dir=r"C:/Users/Lenovo/Desktop"

headers={
'User-Agent': #'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:134.0) Gecko/20100101 Firefox/134.0'
}

'''各学院url'''
# response = requests.get(url=total_url, headers=headers)
# response.encoding = 'utf-8'
# tree = etree.HTML(response.text)

tree = etree.HTML(page_text)

dl_list=tree.xpath('//*[@id="root"]/div/div[3]/div/div/div/div[2]/div/div/div/ul/li')
for a in dl_list:
    sleep(0.5)
    institution=a.xpath('./p/a/text()')[0].strip()
    print(institution)

    dict={
        '高校名称':university, #大学名称
        '学院':institution, #学院名称
        '院系':'', #院系名称
        '职称':'', #学者的职称
        'url':'',
        'xpath':'',
        'xpath_list':'', #学者列表翻页xpath
        '预期采集人数':'',
        '备注':''
    }
    result.append(dict)
    print('--------------------------------------------')

'''保存数据到Excel文件'''
ff = pd.DataFrame(result)
file_path = f"{work_dir}/qs12.xlsx"

bro.close()

if not os.path.exists(file_path):
    # 文件不存在，直接创建新文件
    ff.to_excel(file_path, index=False)
else:
    # 文件存在，读取现有数据并追加新数据
    d1 = pd.read_excel(file_path)
    d1 = pd.concat([d1, ff], ignore_index=True)  # 合并数据
    
    book = load_workbook(file_path)  # 该文件必须存在,并且该语句必须在 with pd.ExcelWriter() 之前
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}
        d1.to_excel(writer, sheet_name="Sheet1", index=False)  # 重写sheet
        