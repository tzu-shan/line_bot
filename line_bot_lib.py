import openpyxl
import my_fun
import datetime

dataWorkbook = openpyxl.load_workbook('miami_line_bot.xlsx')


# create keyword list
def find_message_type(inputText):
    keywordSheet = dataWorkbook['keywords']
    keywords = my_fun.xlsx_get_data_remove_column_name(keywordSheet)
    messageType = ''
    for i in range(len(keywords)):
        if inputText.lower().find(str(keywords[i][0]).lower()) > -1:
            messageType = keywords[i][1]
            break
    return messageType


def reply_driver_info(inputDriver):
    sheet = dataWorkbook['driverName']
    nameList = my_fun.xlsx_get_data(sheet)
    driverSheet = dataWorkbook['driverInfo']
    driverInfo = my_fun.xlsx_get_data_remove_column_name(driverSheet)
    columnName = my_fun.xlsx_get_column_name(driverSheet)
    index = 0
    for i in range(len(nameList)):
        for j in range(len(nameList[i])):
            if inputDriver.lower().find(str(nameList[i][j]).lower()) > -1:
                index = i
                break
    driverInfoReply = []  # 存取車手資訊
    for k in range(1, len(driverInfo[index]) - 1):
        driverInfoReply.append({"type": "box", "layout": "baseline",
                                "contents": [
                                    {"type": "text", "text": columnName[k], "margin": "sm", "flex": 3, "wrap": True,
                                     "align": "end", "size": "sm", "color": "#aaaaaa"},
                                    {"type": "text", "text": str(driverInfo[index][k]), "size": "lg", "wrap": True,
                                     "margin": "md", 'flex': 4}]})
    reply = [{'type': 'flex', 'altText': 'Driver Information',
              'contents': {"type": "bubble",
                           "hero": {"type": "image", "url": driverInfo[index][-1], "size": "full", "aspectMode": "fit"},
                           "body": {"type": "box", "layout": "vertical", "spacing": "md",
                                    "contents": [{"type": "text", "text": driverInfo[index][0], "size": "xl",
                                                  "weight": "bold"},
                                                 {"type": "box", "layout": "vertical", "spacing": "sm",
                                                  "contents": driverInfoReply}]}}}]
    return reply


def reply_team_info():
    teamInfoSheet = dataWorkbook['teamInfo']
    teamInfo = my_fun.xlsx_get_data_remove_column_name(teamInfoSheet)
    teamsInfoReply = []
    for i in range(10):
        teamsInfoReply.append({'type': 'bubble',
                               'header': {'type': 'box', 'layout': 'horizontal',
                                          'contents': [{'type': 'text', 'text': teamInfo[i][0], 'wrap': True}]},
                               'hero': {'type': 'image', 'url': teamInfo[i][2], 'size': 'full'},
                               'body': {'type': 'box', 'layout': 'vertical',
                                        'contents': [{'type': 'text', 'wrap': True,
                                                      'text': f'Base: {teamInfo[i][3]}\n'
                                                              f'World Championships: {teamInfo[i][4]}\n'
                                                              f'Drivers: {teamInfo[i][5]} & {teamInfo[i][6]}'}]},
                               'footer': {'type': 'box', 'layout': 'horizontal',
                                          'contents': [{'type': 'button', 'style': 'link',
                                                        'action': {'type': 'uri', 'label': 'Website',
                                                                   'uri': teamInfo[i][1]}}]}})
    reply = [{'type': 'flex', 'altText': 'Teams Information',
              'contents': {'type': 'carousel', 'contents': teamsInfoReply}}]
    return reply


def reply_stock():
    today = datetime.datetime.now().strftime('%Y%m%d')
    url = f'https://www.twse.com.tw/exchangeReport/BWIBBU_d?response=json&date={today}&selectType=ALL&_=1651891543707'
    file = my_fun.get_json_from_web(url)
    if file['stat'] == 'OK':
        statName = ['殖利率(%)', '本益比', '股價淨值比']
        data = []
        for i in range(len(file['data'])):
            data.append([file['data'][i][j] for j in [0, 1, 2, 4, 5]])
        top10s = [get_top_10_name(2, data), get_top_10_name(3, data), get_top_10_name(4, data)]

        allContent = []
        for a in range(3):
            subContent = [{"type": "text", "text": statName[a], "weight": "bold", "size": "lg"},
                          {"type": "separator", "margin": "lg"}]
            for b in range(10):
                subContent.append({"type": "box", "layout": "horizontal",
                                   "contents": [{"type": "text", "text": top10s[a][b][0], "flex": 2},
                                                {"type": "text", "text": top10s[a][b][1], "flex": 2},
                                                {"type": "text", "text": top10s[a][b][2], "flex": 2, "align": "end"}]})
            allContent.append({"type": "bubble", "size": "kilo",
                                "body": {"type": "box", "layout": "vertical", "contents": subContent, "spacing": "sm",
                                         "paddingAll": "13px"}})
        reply = [{'type': 'text', 'text': f'{today} \n前十殖利率(%)、本益比與股價淨值比'},
                 {'type': 'flex', 'altText': 'Stock Information',
                  'contents': {'type': 'carousel', 'contents': allContent}}]
    else:
        reply = [{'type': 'text', 'text': '很抱歉，今日無交易資訊'}]
    return reply

def get_top_10_name(statIndex, dataset):
    lst = [dataset[i][statIndex] for i in range(len(dataset))]
    temp = lst.copy()
    temp.sort(reverse=True)
    top10Num = temp[0:10]
    top10Index = [lst.index(i) for i in top10Num]
    top10 = []
    for i in range(10):
        top10.append([dataset[top10Index[i]][0], dataset[top10Index[i]][1], dataset[top10Index[i]][statIndex]])
    return top10


def reply_location():
    sheet = dataWorkbook['location']
    data = my_fun.xlsx_get_data(sheet)
    reply = [{'type': 'text', 'text': data[0][1]},
             {'type': 'location', 'title': 'Race Venue',
              'address': data[1][1], 'latitude': data[2][1], 'longitude': data[2][1]}]
    return reply


def reply_schedule():
    sheet = dataWorkbook['schedule']
    data = my_fun.xlsx_get_data_remove_column_name(sheet)
    reply = [{'type': 'flex', 'altText': 'Circuit Infomation',
              'contents': {'type': 'bubble',
                           'body': {'type': 'box', 'layout': 'vertical',
                                    'contents': [
                                        {'type': 'text', 'text': 'SCHEDULE', 'weight': 'bold', 'color': '#1DB446',
                                         'size': 'sm'},
                                        {'type': 'text', 'text': 'MIAMI GP', 'weight': 'bold', 'size': 'xxl',
                                         'margin': 'md'},
                                        {'type': 'text', 'text': 'Miami International Autodrome', 'size': 'xs',
                                         'color': '#aaaaaa', 'wrap': True},
                                        {'type': 'separator', 'margin': 'xxl'},
                                        {'type': 'box', 'layout': 'vertical', 'margin': 'xxl', 'spacing': 'sm',
                                         'contents': [
                                             {'type': 'text', 'text': 'Track Time', 'style': 'italic', 'size': 'xs'},
                                             {'type': 'box', 'layout': 'horizontal',
                                              'contents': [{'type': 'text', 'text': 'Practice 1', 'size': 'sm',
                                                            'color': '#555555', 'flex': 0},
                                                           {'type': 'text', 'text': data[0][1],
                                                            'size': 'sm', 'color': '#111111', 'align': 'end'}]},
                                             {'type': 'box', 'layout': 'horizontal',
                                              'contents': [{'type': 'text', 'text': 'Practice 2', 'size': 'sm',
                                                            'color': '#555555', 'flex': 0},
                                                           {'type': 'text', 'text': data[1][1],
                                                            'size': 'sm', 'color': '#111111', 'align': 'end'}]},
                                             {'type': 'box', 'layout': 'horizontal',
                                              'contents': [
                                                  {'type': 'text', 'text': 'Practice 3', 'size': 'sm',
                                                   'color': '#555555', 'flex': 0},
                                                  {'type': 'text', 'text': data[2][1], 'size': 'sm',
                                                   'color': '#111111', 'align': 'end'}]},
                                             {'type': 'box', 'layout': 'horizontal',
                                              'contents': [
                                                  {'type': 'text', 'text': 'Qualifying', 'size': 'sm',
                                                   'color': '#555555', 'flex': 0},
                                                  {'type': 'text', 'text': data[3][1], 'size': 'sm',
                                                   'color': '#111111', 'align': 'end'}]},
                                             {'type': 'box', 'layout': 'horizontal',
                                              'contents': [
                                                  {'type': 'text', 'text': 'Race', 'size': 'sm', 'color': '#555555',
                                                   'flex': 0},
                                                  {'type': 'text', 'text': data[4][1], 'size': 'sm',
                                                   'color': '#111111', 'align': 'end'}]},
                                             {'type': 'text', 'text': 'Local Time (UTC +08:00)', 'style': 'italic',
                                              'size': 'xs', 'margin': 'xxl'},
                                             {'type': 'box', 'layout': 'horizontal',
                                              'contents': [
                                                  {'type': 'text', 'text': 'Practice 1', 'size': 'sm',
                                                   'color': '#555555', 'flex': 0},
                                                  {'type': 'text', 'text': data[0][2], 'size': 'sm',
                                                   'color': '#111111', 'align': 'end'}]},
                                             {'type': 'box', 'layout': 'horizontal',
                                              'contents': [
                                                  {'type': 'text', 'text': 'Practice 2', 'size': 'sm',
                                                   'color': '#555555', 'flex': 0},
                                                  {'type': 'text', 'text': data[1][2], 'size': 'sm',
                                                   'color': '#111111', 'align': 'end'}]},
                                             {'type': 'box', 'layout': 'horizontal',
                                              'contents': [
                                                  {'type': 'text', 'text': 'Practice 3', 'size': 'sm',
                                                   'color': '#555555', 'flex': 0},
                                                  {'type': 'text', 'text': data[2][2], 'size': 'sm',
                                                   'color': '#111111', 'align': 'end'}]},
                                             {'type': 'box', 'layout': 'horizontal',
                                              'contents': [
                                                  {'type': 'text', 'text': 'Qualifying', 'size': 'sm',
                                                   'color': '#555555', 'flex': 0},
                                                  {'type': 'text', 'text': data[3][2], 'size': 'sm',
                                                   'color': '#111111', 'align': 'end'}]},
                                             {'type': 'box', 'layout': 'horizontal',
                                              'contents': [
                                                  {'type': 'text', 'text': 'Race', 'size': 'sm', 'color': '#555555',
                                                   'flex': 0},
                                                  {'type': 'text', 'text': data[4][2], 'size': 'sm',
                                                   'color': '#111111', 'align': 'end'}]}]}]},
                           'styles': {'footer': {'separator': True}}}}]
    return reply


def reply_info():
    sheet = dataWorkbook['info']
    data = my_fun.xlsx_get_data(sheet)
    reply = [{'type': 'flex', 'altText': 'Circuit Infomation',
              'contents': {'type': 'bubble',
                           'hero': {'type': 'image', 'url': data[3][1], 'size': 'full', 'aspectRatio': '20:13',
                                    'aspectMode': 'fit', 'action': {'type': 'uri', 'uri': data[3][1]}},
                           'body': {'type': 'box', 'layout': 'vertical',
                                    'contents': [
                                        {'type': 'text', 'text': 'Circuit Information', 'weight': 'bold', 'size': 'xl'},
                                        {'type': 'box', 'layout': 'vertical', 'margin': 'lg', 'spacing': 'sm',
                                         'contents': [{'type': 'box', 'layout': 'baseline', 'spacing': 'sm',
                                                       'contents': [{'type': 'text', 'text': 'Number of Laps',
                                                                     'color': '#aaaaaa', 'size': 'md', 'flex': 5,
                                                                     'wrap': True},
                                                                    {'type': 'text', 'text': str(data[0][1]),
                                                                     'wrap': True,
                                                                     'color': '#666666', 'size': 'md', 'flex': 5,
                                                                     'weight': 'bold'}]}]},
                                        {'type': 'box', 'layout': 'baseline', 'spacing': 'sm',
                                         'contents': [{'type': 'text', 'text': 'Circuit Length',
                                                       'color': '#aaaaaa', 'size': 'md', 'flex': 5,
                                                       'wrap': True},
                                                      {'type': 'text', 'text': data[1][1], 'wrap': True,
                                                       'color': '#666666', 'size': 'md', 'flex': 5,
                                                       'weight': 'bold'}]},
                                        {'type': 'box', 'layout': 'baseline', 'spacing': 'sm',
                                         'contents': [{'type': 'text', 'text': 'Race Distance',
                                                       'color': '#aaaaaa', 'size': 'md', 'flex': 5,
                                                       'wrap': True},
                                                      {'type': 'text', 'text': data[2][1], 'wrap': True,
                                                       'color': '#666666', 'size': 'md', 'flex': 5,
                                                       'weight': 'bold'}]}]}}}]
    return reply


def reply_team_standings():
    reply = [{'type': 'image',
              'originalContentUrl': 'https://pbs.twimg.com/media/FRHsDccXIAA3RrT?format=png&name=medium',
              'previewImageUrl': 'https://pbs.twimg.com/media/FRHsDccXIAA3RrT?format=png&name=medium'}]
    return reply


def reply_driver_standings():
    reply = [{'type': 'image',
              'originalContentUrl': 'https://pbs.twimg.com/media/FRHjkCDXIAI962F?format=jpg&name=medium',
              'previewImageUrl': 'https://pbs.twimg.com/media/FRHjkCDXIAI962F?format=jpg&name=medium'}]
    return reply

def reply_response(inputText):
    wb = openpyxl.load_workbook('305-Line問答題.xlsx')
    sheet = wb['工作表1']
    data = my_fun.xlsx_get_data_remove_column_name(sheet)
    replyTxt = ''
    for i in range(len(data)):
        if inputText == data[i][0]:
            replyTxt = data[i][1]
    reply = [{'type': 'text', 'text': replyTxt}]
    return reply


def reply_bus_status(inputText):
    url = 'https://data.tycg.gov.tw/api/v1/rest/datastore/bf55b21a-2b7c-4ede-8048-f75420344aed?format=json&limit=9999'
    file = my_fun.get_json_from_web(url)
    # （0：正常、1：開始、2：結束）
    for i in range(len(file['result']['records'])):
        if inputText.upper().find(file['result']['records'][i]['BusID']) > -1:
            if file['result']['records'][i]['DutyStatus'] == '0':
                status = '正常'
                break
            elif file['result']['records'][i]['DutyStatus'] == '1':
                status = '開始'
                break
            elif file['result']['records'][i]['DutyStatus'] == '2':
                status = '結束'
                break
            else:
                status = '未知或無該公車車號'
                break

    reply = [{'type': 'text', 'text': f'車輛狀態: {status}'}]
    return reply


def get_reply(inputText):
    keywordSheet = dataWorkbook['keywords']
    keywords = my_fun.xlsx_get_data_remove_column_name(keywordSheet)
    messageType = ''
    for i in range(len(keywords)):
        if inputText.lower().find(str(keywords[i][0]).lower()) > -1:
            messageType = keywords[i][1]
            break

    if messageType == 'schedule':
        reply = reply_schedule()
    elif messageType == 'location':
        reply = reply_location()
    elif messageType == 'info':
        reply = reply_info()
    elif messageType == 'teamStandings':
        reply = reply_team_standings()
    elif messageType == 'driverStandings':
        reply = reply_driver_standings()
    elif messageType == 'teams':
        reply = reply_team_info()
    elif messageType == 'response':
        reply = reply_response(inputText)
    elif messageType == 'driverInfo':
        reply = reply_driver_info(inputText)
    elif messageType == 'bus':
        reply = reply_bus_status(inputText)
    elif messageType == 'stock':
        reply = reply_stock()
    else:
        reply = [{'type': 'text',
                  'text': f"We are soo sorry our 3head cannot understand your 5head question.\n"
                          f"But if you want to make some predictions, here's a form for you!"},
                 {'type': 'text', 'text': 'https://forms.gle/9D2obm4g3ciGvBCe6'}]
    return reply
