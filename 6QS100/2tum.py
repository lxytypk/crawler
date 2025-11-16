import re
import pandas as pd
import requests
from lxml import etree
import os
from openpyxl import load_workbook

'''Technical University of Munich'''
total_url='https://www.tum.de/en/'
university='Technical University of Munich'
result=[]
work_dir=r"C:/Users/Lenovo/Desktop"

headers={
'User-Agent': #'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:134.0) Gecko/20100101 Firefox/134.0'
}

'''各学院url'''
response = requests.get(url=total_url, headers=headers)
response.encoding = 'utf-8'
tree = etree.HTML(response.text)

li_list=tree.xpath('//div[@class="flex__auto"]/ul/li')
# print(div_list)
for div in li_list:
    institution=div.xpath('./a/text()')[0].strip()
    print(institution)

    institution_url = div.xpath('./a/@href')[0]
    print(institution_url)

    dict={
        '高校名称':university, #大学名称
        '学院':institution, #学院名称
        '院系':'', #院系名称
        '职称':'', #学者的职称
        'url':institution_url,
        'xpath':'',
        'xpath_list':'', #学者列表翻页xpath
        '预期采集人数':'',
        '备注':''
    }
    result.append(dict)
    print('--------------------------------------------')

'''保存数据到Excel文件'''
ff = pd.DataFrame(result)
file_path = f"{work_dir}/qs100.xlsx"

if not os.path.exists(file_path):
    ff.to_excel(file_path, index=False)
else:
    # 文件存在，先读取旧数据
    d1 = pd.read_excel(file_path)
    d1 = pd.concat([d1, ff], ignore_index=True)

    # 正确加载工作簿并写入
    book = load_workbook(file_path)
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}  # 关键一行，避免文件结构被破坏！
        d1.to_excel(writer, sheet_name="Sheet1", index=False)
        
