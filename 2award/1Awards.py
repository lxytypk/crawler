import re
import pandas as pd
import requests
from lxml import etree
from openpyxl import load_workbook

'''谈家桢生命科学奖获得者'''

'''
2024
https://www.cctanfoundation.org/plus/list.php?tid=60&TotalResult=18&PageNo=1
https://www.cctanfoundation.org/plus/list.php?tid=60&TotalResult=18&PageNo=2

2023
https://www.cctanfoundation.org/plus/list.php?tid=59&TotalResult=18&PageNo=1
https://www.cctanfoundation.org/plus/list.php?tid=59&TotalResult=18&PageNo=2
'''
url='https://www.cctanfoundation.org/plus/list.php?tid=60'
headers={
'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0'
}

result1=[]
result2=[]
work_dir=r"C:\Users\Lenovo\Desktop\lxy\2award"

response=requests.get(url=url,headers=headers)
response.encoding='utf-8'
tree=etree.HTML(response.text)
# print(response.text)

'''年份、url'''
origin_url='https://www.cctanfoundation.org'
d_list=tree.xpath('//div[@class="swiper-wrapper"]/div')
for d in d_list:
    year=d.xpath('./a/text()')[0]
    print(year)
    new =d.xpath('./a/@href')[0]
    # print(new)
    n_url=origin_url+new
    for i in range(1,3):
        new_url=n_url+'&TotalResult=18&PageNo='+str(i)
        # print(new_url)
        b_response=requests.get(url=new_url,headers=headers)
        b_response.encoding = 'utf-8'
        new_tree = etree.HTML(b_response.text)
        div_list=new_tree.xpath('//div[@class="list-box clearfix"]/div')
        for div in div_list:
            '''姓名'''
            name=div.xpath('./a/h5/text()')[0]
            if not name.strip():  # 检查 name 是否为空
                continue  # 跳过该条记录
            '''职称'''
            try:
                profession = div.xpath('./a/h5/span/text()')[0]
                if not profession.strip():  # 检查 profession 是否为空
                    profession = ''  # 设置默认值
            except IndexError:
                profession = ''  # 设置默认值
            '''奖项名称'''
            try:
                reward=div.xpath('./a/p/text()')[0]
                if not reward.strip():  # 检查 profession 是否为空
                    reward = ''  # 设置默认值
            except IndexError:
                reward = ''  # 设置默认值

            # print(reward)
            # print(profession)
            # print(name)
            # print('----------------------')

            dict1={
                '业务所':'生物经济研究所',
                '标签类别（关注的组织机构及人才类别）':'谈家桢生命科学奖获得者',
                '姓名':name,
                '机构':'',
                '奖项名称':reward,
                '获奖类型':'',
                '级别':'',
                '获奖原因':'',
                '获奖项目名称':'',
                '获得时间':year,
                '来源':new_url
            }
            result1.append(dict1)
            dict1={}

            dict2={
                '业务所':'生物经济研究所',
                '标签类别（关注的组织机构及人才类别）':'谈家桢生命科学奖获得者',
                '姓名':name,
                '机构':'',
                '学院':'',
                '性别':'',
                '职称（例如，教授，讲师，长江学者等）':profession,
                '电话':'',
                '邮箱':'',
                '出生日期':'',
                '学历':'',
                '学位':'',
                '研究方向':'',
                '工作职务（例：高等学校教师）':'',
                '个人主页':'',
                '个人简介（个人详情）':''
            }
            result2.append(dict2)
            dict2={}

'''存储文件'''
ff=pd.DataFrame(result1)
file_path = f"{work_dir}/data.xlsx"
ff.to_excel(file_path, sheet_name='奖状字段',index=False,encoding='utf_8')
df=pd.DataFrame(result2)
book = load_workbook(file_path)  # 该文件必须存在,并且该语句必须在 with pd.ExcelWriter() 之前
with pd.ExcelWriter(file_path) as writer:
    writer.book = book
    df.to_excel(writer, sheet_name="采集字段", index=False)