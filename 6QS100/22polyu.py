import re
import pandas as pd
import requests
from lxml import etree
import os
from openpyxl import load_workbook

'''The Hong Kong Polytechnic University'''
total_url='https://www.polyu.edu.hk/sc/education/faculties-schools-departments/'
university='The Hong Kong Polytechnic University'
result=[]
work_dir=r"C:/Users/Lenovo/Desktop"

headers={
'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0',
    #'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:134.0) Gecko/20100101 Firefox/134.0'
}

'''各学院url'''
response = requests.get(url=total_url, headers=headers, verify=False)
response.encoding = 'utf-8'
tree = etree.HTML(response.text)
# print(response.text)

div_list=tree.xpath('//div[@class="item-under-angle-line-blk"]/div[@style="min-height:400px;"]/div[position()>1]')
# print(div_list)
for div in div_list:
    d_list=div.xpath('./div')
    for d in d_list:
        institution=d.xpath('./div[@class="ITS_Content_News_Highlight_Collection  "]/div/div/div/div[1]/p/a/span/text()')[0]
        print(institution)

        li_list=d.xpath('./div/div/div/div/div[1]/ul/li')
        if li_list:
            for li in li_list:
                department=li.xpath('./a/span/text()')[0]
                print(department)
                department_url = li.xpath('./a/@href')[0]
                if not department_url.startswith('http'):
                    department_url='https://www.polyu.edu.hk'+department_url
                print(department_url)

                '''访问每个学院的页面，获取师资力量的URL'''
                institution_response = requests.get(url=department_url, headers=headers, verify=False)
                institution_response.encoding = 'utf-8'
                institution_tree = etree.HTML(institution_response.text)
                try:
                    a_url=institution_tree.xpath('//a[contains(text(), "Our People") or contains(text(), "Our Faculty") or contains(text(), "學者名錄") or contains(text(), "Academic Staff")]/@href')[0]
                except:
                    a_url=''
                # print(a_url)
                url=department_url+a_url
                print(url)

                dict={
                    '高校名称':university, #大学名称
                    '学院':institution, #学院名称
                    '院系':department, #院系名称
                    '职称':'', #学者的职称
                    'url':url,
                    'xpath':'',
                    'xpath_list':'', #学者列表翻页xpath
                    '预期采集人数':'',
                    '备注':''
                }
                result.append(dict)
                print('--------------------------------------------')
        else:
            institution_url = d.xpath('./div/div/div/div/div[1]/p/a/@href')[0]
            if not institution_url.startswith('http'):
                    institution_url='https://www.polyu.edu.hk'+institution_url
            print(institution_url)

            '''访问每个学院的页面，获取师资力量的URL'''
            institution_response = requests.get(url=institution_url, headers=headers, verify=False)
            institution_response.encoding = 'utf-8'
            institution_tree = etree.HTML(institution_response.text)
            try:
                a_url=institution_tree.xpath('//a[contains(text(), "Our People") or contains(text(), "Our Faculty") or contains(text(), "學者名錄") or contains(text(), "Academic Staff")]/@href')[0]
            except:
                a_url=''
            print(a_url)
            url=institution_url+a_url
            print(url)

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
file_path = f"{work_dir}/qs11.xlsx"

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
        