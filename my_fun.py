import xml.etree.ElementTree
import urllib.request as httplib
import ssl
import csv
import matplotlib.pyplot as plt
import json

def get_ubike_data(inputText):
    url = 'https://data.tycg.gov.tw/api/v1/rest/datastore/a1b4714b-3b75-4ff8-a8f2-cc377e4eaa0f?format=json'
    data = get_json_from_web(url)
    allStops = [data['result']['records'][s]['sna'] for s in range(len(data['result']['records']))]
    str1 = ''
    for i in range(len(data['result']['records'])):
        if inputText.find(data['result']['records'][i]['sna']) > -1:
            str1 = f"中文場站名稱: {data['result']['records'][i]['sna']} <br>" \
                   f"可藉數量/總數: {int(data['result']['records'][i]['sbi'])}/{data['result']['records'][i]['tot']} <br>" \
                   f"地址: {data['result']['records'][i]['sarea']}{data['result']['records'][i]['ar']}"
            break
    if str1 == '':
        str1 += '找不到該車站，請再試一次<br><br>站名一覽：<br>'
        str1 += '<br>'.join(allStops)
    return str1

def get_json_from_web(url):
    context = ssl._create_unverified_context()
    with httplib.urlopen(url, context=context) as f:
        dta = json.loads(f.read().decode('utf-8-sig'))
    return dta

def open_csv(fileName):
    dta = []
    with open(fileName, 'r', encoding='utf-8') as f:
        file = csv.reader(f, delimiter=',')
        header = next(file)
        for row in file:
            dta.append(row)
    return header, dta

def save_file(fileName, dataset):
    with open(fileName, 'w+', newline='', encoding='utf-8') as f:
        for line in dataset:
            f.write(line)

def chinese_font():
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
    plt.rcParams['axes.unicode_minus'] = False

def get_url_file(url):
    # 因.urlopen發生問題，將ssl憑證排除
    ssl._create_default_https_context = ssl._create_unverified_context
    url = url
    req = httplib.Request(url, data=None,
                          headers={
                              'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, "
                                            "like Gecko) Chrome/100.0.4896.88 Safari/537.36"})
    reponse = httplib.urlopen(req)  # 開啟連線動作
    contents = []
    if reponse.code == 200:  # 當連線正常時
        contents = reponse.read().decode("utf-8")  # 讀取網頁內容 轉換編碼為 utf-8
    return contents

def open_xml_file_1(name):
    tree = xml.etree.ElementTree.parse(name)
    dta = tree.getroot()
    return dta

def open_xml_file_2(name):
    str = ""
    f = open(name, 'r', encoding='utf-8')
    for line in f:
        str = str + line
    f.close()
    dta = xml.etree.ElementTree.fromstring(str)
    return dta

def xlsx_get_row(sheet, num):
    data = []
    for i in range(1, sheet.max_column+1):
        data.append(sheet.cell(row=num, column=i).value)
    return data

def xlsx_get_col(sheet, num):
    data = []
    for i in range(1, sheet.max_row+1):
        data.append(sheet.cell(row=i, column=num).value)
    return data

def xlsx_get_data(sheet):
    data = []
    for i in range(1, sheet.max_row+1):
        row = []
        for j in range(1, sheet.max_column+1):
            row.append(sheet.cell(row=i, column=j).value)
        data.append(row)
    return data

def xlsx_get_data_remove_column_name(sheet):
    data = []
    for i in range(2, sheet.max_row+1):
        row = []
        for j in range(1, sheet.max_column+1):
            row.append(sheet.cell(row=i, column=j).value)
        data.append(row)
    return data

def xlsx_get_column_name(sheet):
    data = []
    for i in range(1, sheet.max_column+1):
        data.append(sheet.cell(row=1, column=i).value)
    return data

def csv_get_row(fileName, num):
    with open(fileName, 'r', encoding='utf-8') as file:
        i = 0
        dataCSV = csv.reader(file, delimiter=',')
        for row in dataCSV:
            if i == num:
                return row
            i += 1

def csv_get_col(fileName, num, startRow=0):
    data = []
    with open(fileName, 'r', encoding='utf-8') as file:
        i = 0
        dataCSV = csv.reader(file, delimiter=',')
        for row in dataCSV:
            if i >= startRow:
                data.append(row[num])
            i += 1
    return data
