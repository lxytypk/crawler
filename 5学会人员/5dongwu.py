import re
import pandas as pd
import os
import requests
from lxml import etree
from openpyxl import load_workbook

'''中国动物学会'''
total_url='http://czs.ioz.cas.cn/gyxh/xrld/'
university='中国动物学会'
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

def extract_institution(text):
    match = re.search(r'(.*?)(：|: )(.*)时间(.*?)(：|: )(.*)', text)
    if match:
        return match.group(3).strip()
    return None


div_list=tree.xpath('//div[@class="TRS_Editor"]/table')

for d in div_list:
    a_list=d.xpath('./tbody/tr[2]/td')
    for a in a_list:
        try:
            name=a.xpath('.//a/text()')[0].strip()
            print(name)
            job='副理事长'
            print(job)

            try:
                institution_url= a.xpath('./p/a/@href')[0]
                print(institution_url)
                institution_response = requests.get(url=institution_url, headers=headers, verify=False)
                institution_response.encoding = 'utf-8'
                institution_tree = etree.HTML(institution_response.text)

                institution=institution_tree.xpath('/html/body/table[5]/tbody/tr/td/table/tbody/tr[2]/td/text()')[0]
                institution=extract_institution(institution)
                print(institution)
            except:
                institution = ''
                print(institution)

            dict = {
                '学会理事URL': total_url,
                '学会名称': university,
                '姓名': name,
                '职位': job,
                '机构': institution,
                '邮箱': '',
                '任职开始年份': '2024年',
                '任职结束年份': '2029年'
            }
            result.append(dict)
            print('--------------------------------------------')
        except:
            pass

'''续写Excel文件'''
ff=pd.DataFrame(result)
file_path = f"{work_dir}/lxy0314.xlsx"
d1=pd.read_excel(file_path)
d1 = pd.concat([d1, ff], ignore_index=True)	# 合并数据

book = load_workbook(file_path)  # 该文件必须存在,并且该语句必须在 with pd.ExcelWriter() 之前
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    writer.book = book
    d1.to_excel(writer, sheet_name="Sheet1",index=False)  # 重写sheet