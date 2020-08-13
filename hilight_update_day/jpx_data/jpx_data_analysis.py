# JPX取得
import csv
import datetime
import json
import configparser
import urllib.error
import time
import os
import urllib.request
import re
import pprint
import xlrd
import requests
import errno
from bs4 import BeautifulSoup

print("--------------------------------------------------------------------------------------------")
print("config設定")
print("--------------------------------------------------------------------------------------------")
config_ini = configparser.ConfigParser()
config_ini_name = 'config.ini'
config_dir = os.path.dirname(__file__)
config_ini_path = os.path.join(config_dir, config_ini_name)

print("config file path")
print(config_ini_path)

print("config読み込み")
if not os.path.exists(config_ini_path):
    raise FileNotFoundError(errno.ENOENT, os.strerror(
        errno.ENOENT), config_ini_path)
config_ini.read(config_ini_path, encoding='utf-8')
target_code = config_ini.get('Brandcode', 'targets')
target_code = json.loads(target_code)

print("--------------------------------------------------------------------------------------------")
print("datetimeへ置換")
print("--------------------------------------------------------------------------------------------")
def datetime_replace(num):
    from datetime import datetime, timedelta
    return(datetime(1899, 12, 30) + timedelta(days=num)).strftime("%Y/%m/%d")

print("--------------------------------------------------------------------------------------------")
print("result output txt")
print("--------------------------------------------------------------------------------------------")
def result_output(macth_list):
    #実行日
    execute_date = ''
    result_file_name = '_result.csv'
    dt_now = datetime.datetime.now()
    execute_date = dt_now.strftime('%Y%m%d_%H%M%S')
    result_file_name = execute_date + result_file_name

    #実行結果
    result_data = macth_list

    #新規作成・書き込み
    with open(result_file_name, 'a') as file:
        writer = csv.writer(file, lineterminator='\n')
        writer.writerow(result_data)


print("--------------------------------------------------------------------------------------------")
print("シート内データ解析開始")
print("--------------------------------------------------------------------------------------------")
def analize_sheet_data(wb):
    print("エクセルのじーと名")
    print(wb.sheet_names())
    print("シート0の行列")
    sheet_1 = wb.sheet_by_index(0)
    print(sheet_1.ncols)
    print(sheet_1.nrows)

    #行
    for col in range(sheet_1.ncols):
        #以下の前提
        # 決算発表日 row 0 float →str datetime 未定という文字列が入る場合は無視
        # コード     row 1 float →intへキャスト
        # 会社名     row 2 string →なし
        # 決算期末   row 3 string →なし
        # 業種名     row 4 string →なし
        # 種別       row 5 string →なし
        # 市場区分   row 6 string →なし
        #最初の3行がセル結合されていて面倒なのでプラスしてください
        print('----------------------------')

        for row in range(sheet_1.nrows-5):
            # 0 1 のみ処理が必要2 3 4 5 6は不要
            if col == 0:
                #未定の文字列の場合がある
                #print(type(sheet_1.cell(row+3, col).value))
                if type(sheet_1.cell(row+3, col).value) == float:
                    print(datetime_replace(sheet_1.cell(row+3, col).value))
                else:
                    print(sheet_1.cell(row+3, col).value)
            elif col == 1:
                #最初の3行がセル結合されていて面倒なのでプラスしてください
                #取得時floatのためintにします
                code = int(sheet_1.cell(row+3, col).value)
                #print(int(sheet_1.cell(row+3, col).value))
                #configの銘柄と一致するか検索
                if code in target_code:
                    print('----------------------------macth----------------------------')
                    print(datetime_replace(sheet_1.cell(row+3, 0).value))
                    print(code)
                    print(sheet_1.cell(row+3, 2).value)
                    print(sheet_1.cell(row+3, 3).value)
                    print(sheet_1.cell(row+3, 4).value)
                    print(sheet_1.cell(row+3, 5).value)
                    print(sheet_1.cell(row+3, 6).value)
                    macth_list = []
                    macth_list.append(datetime_replace(sheet_1.cell(row+3, 0).value))
                    macth_list.append(str(code))
                    macth_list.append(sheet_1.cell(row+3, 2).value)
                    macth_list.append(sheet_1.cell(row+3, 3).value)
                    macth_list.append(sheet_1.cell(row+3, 4).value)
                    macth_list.append(sheet_1.cell(row+3, 5).value)
                    macth_list.append(sheet_1.cell(row+3, 6).value)
                    print('----------------------------macth----------------------------')
                    print('----------------------------result書き込み--------------------')
                    print(macth_list)
                    result_output(macth_list)
                else:
                    print('not macth')


print("--------------------------------------------------------------------------------------------")
print("dataフォルダ確認")
print("--------------------------------------------------------------------------------------------")
# カレント
print(os.path.dirname(os.path.abspath("__file__")))
c_dir = os.path.dirname(os.path.abspath("__file__"))

# ダウンロードフォルダ
dl_folder_name = 'data'

#ptah join
dirPath = os.path.join(c_dir, dl_folder_name)
print(dirPath)
result = [f for f in os.listdir(
    dirPath) if os.path.isfile(os.path.join(dirPath, f))]
print(result)

print("--------------------------------------------------------------------------------------------")
print("エクセルライブラリ")
print("--------------------------------------------------------------------------------------------")
for g in result:
    print(g)
    wb = xlrd.open_workbook(os.path.join(dirPath, g))
    analize_sheet_data(wb)
