import re
import pandas as pd
import os
import requests
from lxml import etree
from openpyxl import load_workbook

'''中国力学学会'''
total_url='https://www.cstam.org.cn/getLjMemberList/7339/1'
university='中国力学学会'
result=[]
work_dir=r"C:\Users\Lenovo\Desktop\0314"

headers={
'User-Agent': #'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:134.0) Gecko/20100101 Firefox/134.0'
}

'''各学院url'''
response = requests.get(url=total_url, headers=headers)
response.encoding = 'utf-8'
tree = etree.HTML(response.text)

def extract_text_after_colon(text):
    match = re.search(r'(.*?)：(.+)(院士|教授|研究员)', text)
    if match:
        return match.group(2).strip()
    return None

def extract_job(text):
    match = re.search(r'(.*?)：(.+)(院士|教授|研究员)', text)
    if match:
        return match.group(1).strip()
    return None

def extract_job2(text):
    match = re.search(r'(.*?)：', text)
    if match:
        return match.group(1).strip()
    return None

def extract_institution(text):
    match = re.search(r'(男|女)，(.*?)(大学|研究所|研究院)', text)
    if match:
        return match.group(2)+match.group(3).strip()
    return None

# div_list=tree.xpath('//div[@class="mid-right"]/div/ul')
#
# for d in div_list:
#     li_list=d.xpath('./li')
#     for li in li_list:
#         name00=li.xpath('./a/span/text()')[0]
#         name=extract_text_after_colon(name00)
#         print(name)
#         job = extract_job(name00)
#         print(job)
#
#         institution_url = li.xpath('./a/@href')[0]
#         if not institution_url.startswith('http'):
#             institution_url='https://www.cstam.org.cn'+institution_url
#         institution_response = requests.get(url=institution_url, headers=headers, verify=False)
#         institution_response.encoding = 'utf-8'
#         institution_tree = etree.HTML(institution_response.text)
#
#         try:
#             txt=institution_tree.xpath('//div[@class="txt-wp"]/p[@style][1]//text()')[0]
#             institution=extract_institution(txt)
#             print(institution)
#         except:
#             institution=''
#
#         dict = {
#             '学会理事URL': total_url,
#             '学会名称': university,
#             '姓名': name,
#             '职位': job,
#             '机构': institution,
#             '邮箱': '',
#             '任职开始年份': '2025年2月15日',
#             '任职结束年份': ''
#         }
#         result.append(dict)
#         print('--------------------------------------------')

div_list=tree.xpath('//div[@class="mid-right"]/div/div')

for d in div_list:
    job = d.xpath('./h3/text()')[0]
    job=extract_job2(job)
    print(job)

    li_list=d.xpath('./div/a')
    for li in li_list:
        name=li.xpath('./text()')[0].strip()
        print(name)
        try:
            institution_url = li.xpath('./@href')[0]
            if not institution_url.startswith('http'):
                institution_url='https://www.cstam.org.cn'+institution_url
            institution_response = requests.get(url=institution_url, headers=headers, verify=False)
            institution_response.encoding = 'utf-8'
            institution_tree = etree.HTML(institution_response.text)

            txt=institution_tree.xpath('//div[@class="txt-wp"]/p[@style][1]//text()')[0]
            institution=extract_institution(txt)
            print(institution)
        except:
            institution=''

        dict = {
            '学会理事URL': total_url,
            '学会名称': university,
            '姓名': name,
            '职位': job,
            '机构': institution,
            '邮箱': '',
            '任职开始年份': '2025年2月15日',
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