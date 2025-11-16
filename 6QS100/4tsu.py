import re
import pandas as pd
import requests
from lxml import etree
import os
from openpyxl import load_workbook

'''The Australian National University'''
total_url='https://www.anu.edu.au/about/academic-colleges'
university='The Australian National University'
result=[]
work_dir=r"F:\save\lxy__\pachong\lxy\8QS100"

headers={
'User-Agent': #'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:134.0) Gecko/20100101 Firefox/134.0'
}

'''各学院url'''
response = requests.get(url=total_url, headers=headers)
response.encoding = 'utf-8'
tree = etree.HTML(response.text)

dl_list=tree.xpath('//div[@class="orgaCon"]/dl')
for dl in dl_list:
    a_list=dl.xpath('./div/dd')
    # print(a_list)
    for a in a_list:
        institution=a.xpath('./h3/a/text() | ./h4/a/text()')[0]
        print(institution)
        if(institution=='* 理学院 * '):
            institution_url=''
            continue
        else:
            institution_url = a.xpath('./h3/a/@href | ./h4/a/@href')[0]
        print(institution_url)

        '''访问每个学院的页面，获取师资力量的URL'''
        institution_response = requests.get(url=institution_url, headers=headers, verify=False)
        institution_response.encoding = 'utf-8'
        institution_tree = etree.HTML(institution_response.text)
        try:
            a_url=institution_tree.xpath('//a[contains(text(), "师资") or contains(text(), "教师") or contains(text(), "教职员工") or contains(text(), "学术团队")]/@href')[0]
        except:
            a_url=''
        print(a_url)
        url=institution_url+a_url
        print(url)

        dict={
            'url':url,
            'xpath':'',
            'xpath_list':'', #学者列表翻页xpath
            '备注':'',
            '机构':university, #大学名称
            '学院':institution, #学院名称
            '预期title':'', #学者的职称
            '预期采集人数':'',
            '院系':'' #院系名称
        }
        result.append(dict)
        print('--------------------------------------------')

'''续写Excel文件'''
ff=pd.DataFrame(result)
file_path = f"{work_dir}/qs500.xlsx"
d1=pd.read_excel(file_path)
d1 = pd.concat([d1, ff], ignore_index=True)	# 合并数据

book = load_workbook(file_path)  # 该文件必须存在,并且该语句必须在 with pd.ExcelWriter() 之前
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    writer.book = book
    d1.to_excel(writer, sheet_name="Sheet1",index=False)  # 重写sheet