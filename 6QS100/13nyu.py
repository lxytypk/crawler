import re
import pandas as pd
import requests
from lxml import etree
import os
from openpyxl import load_workbook

'''New York University'''
total_url='https://www.nyu.edu/academics/schools-and-colleges.html'
university='New York University'
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


dl_list=tree.xpath('//section[@class="content"]/div[position()>2]')
for dl in dl_list:
    li_list=dl.xpath('./div/div/ul/li')
    for a in li_list:
        institution=a.xpath('./a/text()')[0].strip()
        print(institution)
        
        institution_url = a.xpath('./a/@href')[0]
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
file_path = f"{work_dir}/qs12.xlsx"

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
        