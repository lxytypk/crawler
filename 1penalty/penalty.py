import requests
from lxml import etree
import re
import openpyxl
import os
import datetime
import pandas as pd
from openpyxl.reader.excel import load_workbook

url='https://www.nsfc.gov.cn/publish/portal0/jd/04/info93663.htm'
headers= {
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0',
    'accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'accept-language':'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    'host':'www.nsfc.gov.cn',
    'referer':'https://www.nsfc.gov.cn/publish/portal0/jd/04/module2248/page1.htm'
}
result=[]
work_dir=r"C:\Users\Lenovo\Desktop\lxy\1penalty"

requests.packages.urllib3.disable_warnings()
page = requests.get(url=url, headers=headers, verify=False)
page.encoding = 'utf-8'
tree = etree.HTML(page.text)
# print(page_text)
p_list=tree.xpath('//*[@id="zoom"]/p')

# 正则表达式，匹配以（a）开始的段落
pattern = r'（[一二三四五六七八九十]+）'  # 匹配类似（一）、（二）等标记
content_list = []
current_content = []
inside_section = False

# 遍历所有的<p>标签内容
for p in p_list:
    text = p.text.strip()
    # print(text)
    # 如果段落内容匹配以（a）开始的格式
    if re.match(pattern, text):
        # 如果当前内容不为空，则将其加入内容列表
        if inside_section:
            content_list.append("\n".join(current_content).strip())
            current_content = []
        inside_section = True
        continue  # 跳过标记本身
    if inside_section:
        current_content.append(text)

# 添加最后一段内容
if current_content:
    content_list.append("\n".join(current_content))

now=datetime.datetime.now()

for content in content_list:
    '''机构名称'''
    # print(content.split("\n")[0])
    str1=content.split("\n")[0]
    school_pattern = re.compile(r'对(.*?)?涉嫌学术不端')
    match = school_pattern.search(str1)
    if(match):
        school_name = match.group(1)
        pattern1 = re.compile(r'.*?(高校|大学|院|研究所|公司)')
        match1 = pattern1.search(school_name)
        school_name1 = match1.group(0)
    else:
        school_name1 = ''
    # print(school_name1)

    '''处罚项目名称'''
    # print(content.split("\n")[1])
    str_name = content.split("\n")[1]
    name_pattern = re.compile(r'\.(.*?)(\.|。)')
    namematch = name_pattern.search(str_name)
    if (namematch):
        article_name = namematch.group(1)
    else:
        article_name = ''
    # print(article_name)

    '''处罚类型描述'''
    # print(content.split("\n")[-1])
    str_ban = content.split("\n")[-1]
    ban_pattern = re.compile(r'(永?久?取消).*?(项目.*?资格)')
    banmatch = ban_pattern.search(str_ban)
    if (banmatch):
        ban = banmatch.group(1)+banmatch.group(2)
    else:
        ban = ''
    # print(ban)

    '''发布时间、结束时间'''
    # print(content.split("\n")[-1])
    str_time= content.split("\n")[-1]
    time_pattern = re.compile(r'（(\d{4}年\d{1,2}月\d{1,2}日)至(\d{4}年\d{1,2}月\d{1,2}日)）')
    timematch = time_pattern.search(str_time)
    if(timematch):
        start_time = timematch.group(1)
        end_time = timematch.group(2)
        # print(start_time)
        # print(end_time)
    else:
        start_time = ''
        end_time = ''

    dict={
        '人才id':'',
        '处罚名称':'学术不端',
        '处罚项目名称':article_name,
        '处罚类型描述':ban,
        '处罚原因描述':'涉嫌学术不端开展了调查',
        '内容':content,
        '发布时间':start_time,
        '结束时间':end_time,
        '机构名称':school_name1,
        '机构id':'',
        '原文链接地址':url,
        '执行机构名称':'国家自然科学基金委员会监督委员会',
        '数据来源':'国家自然科学基金委员会',
        '创建时间':now.strftime("%Y-%m-%d"),
        '数据更新时间':''
    }
    # print(content)
    result.append(dict)
    dict={}
    print("-" * 30)
    # print("\n")

'''存储文件'''
ff=pd.DataFrame(result)
file_path = f"{work_dir}/data.xlsx"
# ff.to_excel(file_path, sheet_name='处罚信息表',index=False,encoding='utf_8')

if not os.path.exists(file_path):
    ff.to_excel(file_path,  sheet_name='处罚信息表', index=False,encoding='utf_8')
