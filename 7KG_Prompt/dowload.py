import os
import re
import requests
import time
import random
import urllib3
from selenium import webdriver
from lxml import etree
from selenium.webdriver.firefox.options import Options
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
import urllib3.contrib.pyopenssl
urllib3.contrib.pyopenssl.inject_into_urllib3()

def has_chinese(s):
    return re.search('[\u4e00-\u9fa5]',s) is not None

def get_datasets():
    headers = {
        "x-api-key": "fTOyvkpIp77Me5ejBhxSD8BGWhhxMZIXaTd5Nf9v",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36 Edg/141.0.0.0"
    }
    query_list=[
        '飞行器+气动力',
        '飞行器+气动热',
        '飞行器+气动热+不确定性',
        '飞行器+气动力+不确定性',
        '气动力+热特性'
        ]
    base_url='https://api.semanticscholar.org/graph/v1/paper/search'
    limit=100
    url_set=set()
    
    for query in query_list:
        print(f"\n----------------开始搜索：{query}----------------")
        offset=0
        while True:
            params={
                "query":query,
                "fields":"title,openAccessPdf",
                "limit":limit,
                "offset":offset,
                "fieldsOfStudy":"Engineering,Physics,Computer Science"
            }
            response=requests.get(base_url,headers=headers,params=params,verify=False).json()
            # print(response)
            time.sleep(3)
            if 'data' in response:
                for item in response['data']:
                    pdf_info = item.get("openAccessPdf")
                    if pdf_info and pdf_info.get('url'):
                        title=item['title']
                        pdf_url=pdf_info['url']
                        if has_chinese(title):
                            url_set.add((title,pdf_url))
                        
            if 'next' in response and response['next']!=None:
                offset=response['next']
                time.sleep(3)
            else:
                break
        print('-------------------------------------------')
        time.sleep(random.uniform(3, 6))
            
    # for title, pdf_url in url_set:
    #     print(title)
    return url_set

def sanitize_filename(name: str) -> str:
    """清理文件名中的非法字符"""
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    name = name.strip().replace('\n', ' ')
    return name

def find_final_pdf_url(start_url):
    options = Options()
    options.add_argument("--headless")  # 无界面模式

    # 启动 Firefox WebDriver
    driver = webdriver.Firefox(options=options)

    try:
        print(f"访问页面：{start_url}")
        driver.set_page_load_timeout(30) #限制加载超时时间
        driver.get(start_url)
        time.sleep(2)
        final_url = driver.current_url
        print(f"最终URL：{final_url}")
        return final_url
    except Exception as e:
        return start_url
    finally:
        driver.quit()
    
def download_pdfs(url_list,save_dir="papers_chinese"):
    os.makedirs(save_dir,exist_ok=True)
    headers = {
        "x-api-key": "fTOyvkpIp77Me5ejBhxSD8BGWhhxMZIXaTd5Nf9v",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36 Edg/141.0.0.0",
        'Connection':'close'
    }

    for i,(title,url) in enumerate(url_list):
        if not url:
            continue
        try:
            title=sanitize_filename(title)
            print(f"[{i+1}/{len(url_list)}] 正在下载：{url}")
            
            try:
                head_resp=requests.head(url,headers=headers,timeout=10,verify=False,allow_redirects=True)
                content_type=head_resp.headers.get("Content-Type","").lower()
            except:
                content_type=""
            if url.lower().endswith(".pdf"):
                final_url = url
            elif "application/pdf" in content_type:
                final_url = head_resp.url
            else:
                final_url = find_final_pdf_url(url)
            
            response=requests.get(final_url,headers=headers,timeout=90,verify=False)
            content_type=response.headers.get("Content-Type", "").lower()
            if not final_url.lower().endswith('.pdf') or 'application/pdf' not in content_type:
                continue
            file_path=os.path.join(save_dir,f"{title}.pdf")
            with open(file_path,"wb") as f:
                f.write(response.content)
            print(f"已保存到：{file_path}")
            
        except requests.exceptions.SSLError as e:
            print(f"??SSL错误：{url} -> {e}")
        except Exception as e:
            print(f"??下载出错：{url} -> {e}")
        time.sleep(3)
    print("所有文件下载完成!")
    
if __name__ == "__main__":
    url_list=get_datasets()
    download_pdfs(url_list)