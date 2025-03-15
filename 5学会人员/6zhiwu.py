import re
import pandas as pd
import os
import requests
from lxml import etree
from openpyxl import load_workbook

'''中国植物学会'''
total_url='http://www.botany.org.cn/xhjj/xhld/'
university='中国植物学会'
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

def extract_name(text):
    match = re.search(r'(.*?)(院士|教授|研究员)', text)
    if match:
        return match.group(1).strip()
    return None


div_list=tree.xpath('//div[@class="neirong"]/div')


for d in div_list:
    job = d.xpath('./p//strong/text()')[0]
    a=d.xpath('./table/tbody/tr')
    for bb in a:
        c=bb.xpath('./td')
        for dd in c:
            name00=dd.xpath('./span[1]/a//strong/span/strong/text()')[0]
            name=extract_name(name00)
            print(name)

            institution=dd.xpath('./span[2]/p[1]/span/text()')[0].strip()
            print(institution)

            dict = {
                '学会理事URL': total_url,
                '学会名称': university,
                '姓名': name,
                '职位': job,
                '机构': institution,
                '邮箱': '',
                '任职开始年份': '2023年10月',
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