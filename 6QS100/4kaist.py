import re
import pandas as pd
import requests
from lxml import etree
import os
from openpyxl import load_workbook

'''Korea Advanced Institute of Science & Technology'''
total_url='https://www.kaist.ac.kr/en/html/edu/03.html#0309'
university='Korea Advanced Institute of Science & Technology'
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


nn1=tree.xpath('//div[@id="txt"]/div[1]/div[2]/h3/text()')[0].strip()
name1_list=tree.xpath('//div[@id="txt"]/div[1]/div[2]/ul/li')
for name1 in name1_list:
    institution=name1.xpath('./a//text()')[0].strip()
    if not institution.startswith('Department of') and not institution.startswith('Graduate School of'):
        institution = 'Department of ' + institution
    print(institution)
    institution_url = name1.xpath('./a/@href')[0]
    print(institution_url)
    dict={
        '高校名称':university, #大学名称
        '学院':nn1, #学院名称
        '院系':institution, #院系名称
        '职称':'', #学者的职称
        'url':institution_url,
        'xpath':'',
        'xpath_list':'', #学者列表翻页xpath
        '预期采集人数':'',
        '备注':''
    }
    result.append(dict)

name_list=tree.xpath('//div[@id="txt"]')
for name in name_list:
    nn=name.xpath('./h3/text()')[0].strip()
    print(nn)
    institution_list=name.xpath('./ul')
    for a in institution_list:
        c_list=a.xpath('./li')
        if len(c_list) > 1:
            for c in c_list:
                institution=c.xpath('./a//text()')[0].strip()
                if not institution.startswith('Department of') and not institution.startswith('Graduate School of') and not institution.startswith('School of'):
                    institution = 'Department of ' + institution
                print(institution)
                
                institution_url = c.xpath('./a/@href')[0]
                print(institution_url)

                dict={
                    '高校名称':university, #大学名称
                    '学院':nn, #学院名称
                    '院系':institution, #院系名称
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
        