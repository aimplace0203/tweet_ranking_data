import os
import re
import csv
import json
import datetime
import requests
import gspread
from time import sleep
from oauth2client.service_account import ServiceAccountCredentials

# Logger setting
from logging import getLogger, FileHandler, DEBUG
logger = getLogger(__name__)
today = datetime.datetime.now()
os.makedirs('./log', exist_ok=True)
handler = FileHandler(f'log/{today.strftime("%Y-%m-%d")}_result.log', mode='a')
handler.setLevel(DEBUG)
logger.setLevel(DEBUG)
logger.addHandler(handler)
logger.propagate = False

### functions ###
def getLatestDownloadedDirName(downloadsDirPath):
    if len(os.listdir(downloadsDirPath)) == 0:
        return None
    return max (
        [downloadsDirPath + '/' + f for f in os.listdir(downloadsDirPath)],
        key=os.path.getctime
    )

def sendChatworkNotification(message):
    try:
        url = f'https://api.chatwork.com/v2/rooms/{os.environ["CHATWORK_ROOM_ID"]}/messages'
        headers = { 'X-ChatWorkToken': os.environ["CHATWORK_API_TOKEN"] }
        params = { 'body': message }
        requests.post(url, headers=headers, params=params)
    except Exception as err:
        logger.error(f'Error: sendChatworkNotification: {err}')
        exit(1)

def num2alpha(num):
    if num<=26:
        return chr(64+num)
    elif num%26==0:
        return num2alpha(num//26-1)+chr(90)
    else:
        return num2alpha(num//26)+chr(64+num%26)

### Google ###
def getRankingCsvData(csvPath):
    with open(csvPath, newline='', encoding='utf-8') as csvfile:
        buf = csv.reader(csvfile, delimiter=',', lineterminator='\r\n', skipinitialspace=True)
        next(buf)
        for row in buf:
            yield row

def checkUploadData(datas):
    try:
        SPREADSHEET_ID = os.environ['RANK_DATA_SSID']
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
        gc = gspread.authorize(credentials)
        sheet = gc.open_by_key(SPREADSHEET_ID).worksheet('Rank Data')

        message = "[info][title]【事前確認用】本日のPeoPle'sランキングツイート[/title]"
        message += f"ꉂꉂ📢PeoPle's 検索順位速報✨\n\n"
        message += f"◎計測地域：新宿🏙\n"
        message += f"ーー\n"

        for data in datas:
            rdate = datetime.datetime.strptime(data[9], '%b %d, %Y').strftime('%Y/%m/%d')
            if rdate != today.strftime('%Y/%m/%d'):
                message = "[info][title]【事前確認用】本日のPeoPle'sランキングツイート[/title]"
                message += "過去のデータが取得されました。\n担当者は本日の順位計測に問題がないかご確認ください。[/info]"
                sendChatworkNotification(message)
                logger.debug(f'checkUploadData: 過去のデータが取得されました。')
                exit(0)
            keyword = data[0]
            try:
                rank = int(data[3])
                medal = '🏅'
            except Exception as err:
                rank = '-'
                medal = ''
            try:
                diff = int(data[6].replace(' ', ''))
                if diff == 0:
                    arrow = '➡️'
                elif diff > 0:
                    arrow = '↗️'
                else:
                    arrow = '↘️'
            except Exception as err:
                diff = '-'
                arrow = ''
            sheet.append_row([keyword, rank, diff, rdate], value_input_option='USER_ENTERED')
            
            message += f'『{keyword}』 {rank}位{medal}{arrow}\n'

        last_column_num = len(list(sheet.row_values(1)))
        last_column_alp = num2alpha(last_column_num)
        sheet.set_basic_filter(name=(f'A:{last_column_alp}'))

        message += "\n＼check／✌🏻\n"
        message += "https://aimplace.co.jp/p"
        message += '[/info]'
        sendChatworkNotification(message)
    except Exception as err:
        logger.debug(f'Error: checkUploadData: {err}')
        exit(1)

### main_script ###
if __name__ == '__main__':

    try:
        rankDataDirPath = os.environ["RANK_DATA_DIR"]
        dateDirPath = getLatestDownloadedDirName(rankDataDirPath)

        data = list(getRankingCsvData(f'{dateDirPath}/aimplace.co.jp.txt'))
        logger.info(f'ranking: {data}')
        checkUploadData(data)

        logger.info("check_ranking_data: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'check_ranking_data: {err}')
        exit(1)
