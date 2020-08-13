# JPX取得
import urllib.error
import time
import os
import urllib.request
import re
import pprint
import xlrd
import requests
import urllib.parse
import ntpath
from bs4 import BeautifulSoup

#sq
r = requests.get("https://www.jpx.co.jp/listing/event-schedules/financial-announcement/index.html")
soup = BeautifulSoup(r.content, "html.parser")


#print("JPX a要素をすべて探し、hrefの値をとる")
# ページに含まれるリンクを全て取得する
links = [url.get('href') for url in soup.find_all('a')]
excel = 'xls'
excel_list = []
links = [e for e in links if e is not None]
excel_list = [l for l in links if excel in l]

print("ヒットしたファイル")
print(excel_list)
print("ヒットしたファイル 個別")
[print(i) for i in excel_list]
print("ヒットしたファイル数")
len(excel_list)


#print("エクセルダウンロード処理")
def download_file(url, dst_path):
    try:
        with urllib.request.urlopen(url) as web_file:
            data = web_file.read()
            with open(dst_path, mode='wb') as local_file:
                local_file.write(data)
    except urllib.error.URLError as e:
        print(e)


def download_file_to_dir(url, dst_dir):

    #data_dir = os.path.join(dst_dir, os.path.basename(url))
    #process_dir = os.path.dirname(os.path.abspath("__file__"))
    
    data_dir = ntpath.join(dst_dir, ntpath.basename(url))
    process_dir = ntpath.dirname(ntpath.abspath("__file__"))
    
    print('パス情報')
    print(data_dir)
    print(process_dir)

    #download_file(url, os.path.join(process_dir, data_dir))

    download_file(url, ntpath.join(process_dir, data_dir))


#print("JPXドメイン（固定）")
jpx_domain = 'https://www.jpx.co.jp/'

#print("ダウンロードフォルダ")
dst_dir = 'data'

#print("エクセルダウンロードパス作成&ダウンロード")
jpx_domain = 'https://www.jpx.co.jp/'

for g in excel_list:
    g = urllib.parse.urljoin(jpx_domain, g)
    print('URL結合')
    print(g)
    print('ダウンロード開始')
    download_file_to_dir(g, dst_dir)
    print("ダウンロード終了")

os.system('pause')
