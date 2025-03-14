import numpy as np
import datetime
import openpyxl
import requests
from lxml import etree
wb = openpyxl.load_workbook("pure_collect.xlsx")
ws = wb.active
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
    'Referer': 'https://www.example.com',
    'Connection': 'keep-alive',
}

'''change'''
Institute="Aalborg University"

xpath_path='//*[@id="main-content"]/div/div[2]/ul/li'
current_date = datetime.date.today()
num=0
for i in range(3):

    '''change'''
    if i ==0:
        website = "https://vbn.aau.dk/en/persons"
        xpath_list=None#'//*[@id="main-content"]/div/div[2]/nav/ul/li[3]/a'
    else:
        website=f"https://vbn.aau.dk/en/persons/?page={i}"
        xpath_list='//*[@id="main-content"]/div/div[2]/nav/ul/li[3]/a'
    response=requests.get(website,headers=headers)
    html=response.content
    tree=etree.HTML(html)
    researcher_list=tree.xpath(xpath_path)
    reseacher_len=len(researcher_list)
    print(reseacher_len)
    num+=reseacher_len  
    

    for i in range(reseacher_len):
        if len(tree.xpath(f'//*[@id="main-content"]/div/div[2]/ul/li[{i+1}]/div/div/ul[2]/li/a'))!=0:
            # print(tree.xpath(f'//*[@id="main-content"]/div/div[2]/ul/li[{i+1}]/div/div/ul[2]/li/a')[0].get('href').split('/'))
            b=tree.xpath(f'//*[@id="main-content"]/div/div[2]/ul/li[{i+1}]/div/div/ul[2]/li/a')[0].get('href')
            print('https://vbn.aau.dk'+b)
            college=('https://vbn.aau.dk'+b).split('/')[5]
            print(college)
        else:
            college=None
        if len(tree.xpath(f'//*[@id="main-content"]/div/div[2]/ul/li[{i+1}]/div/div/h3/a/span'))!=0:
            name=tree.xpath(f'//*[@id="main-content"]/div/div[2]/ul/li[{i+1}]/div/div/h3/a/span')[0].text
            print(name)
        else:
            name=None
        if len(tree.xpath(f'//*[@id="main-content"]/div/div[2]/ul/li[{i+1}]/div/div/h3/a')):
            title=tree.xpath(f'//*[@id="main-content"]/div/div[2]/ul/li[{i+1}]/div/div/h3/a')[0].get('class')
            print(title)
        else:
            title=None
        link='https://vbn.aau.dk'+tree.xpath(f'//*[@id="main-content"]/div/div[2]/ul/li[{i+1}]/div/div/h3/a')[0].get('href')
        print(link)
        person_response=requests.get(link,headers=headers)
        html=person_response.content
        person_tree=etree.HTML(html)
        person_pub_list=person_tree.xpath(f'//*[@class="page-section content-relation-section person-publications"]/div/div[2]/ul/li/div[1]/div[1]/h3/a/span')
        print(person_pub_list)
        for (i,pub) in enumerate(person_pub_list):
            if i ==0:
                data=[1,
                    '2025/1/2',
                    website,
                    xpath_path,
                    xpath_list,
                    Institute,
                    name,
                    title,
                    college,
                    pub.text     
                ]
            else:
                data=[None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    pub.text
                ]  
            ws.append(data)
        if len(person_pub_list)==0:
            data=[1,
                    '2025/1/2',
                    website,
                    xpath_path,
                    xpath_list,
                    Institute,
                    name,
                    title,
                    college,
                    None     
                ]
            ws.append(data)
            
        
            
    
data0=[num,
        '2025/1/10',
        website,
        xpath_path,
        None,
        Institute,
        None,
        None,
        None
            ]
ws.append(data0)
wb.save("pure_collect.xlsx")       

response.close()
