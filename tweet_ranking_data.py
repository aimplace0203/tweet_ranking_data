import os
import re
import csv
import json
import datetime
import requests
from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from fake_useragent import UserAgent
from webdriver_manager.chrome import ChromeDriverManager

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

### Google ###
def getRankingCsvData(csvPath):
    with open(csvPath, newline='', encoding='utf-8') as csvfile:
        buf = csv.reader(csvfile, delimiter=',', lineterminator='\r\n', skipinitialspace=True)
        next(buf)
        for row in buf:
            yield row

def getUploadData(datas):
    try:
        message = f"ꉂꉂ📢PeoPle's 検索順位速報✨%0a%0a"
        message += f"◎計測地域：新宿🏙%0a"
        message += f"ーーー%0a"

        for data in datas:
            rdate = datetime.datetime.strptime(data[9], '%b %d, %Y').strftime('%Y/%m/%d')
            if rdate != today.strftime('%Y/%m/%d'):
                message = "[info][title]PeoPle'sランキングツイート結果[/title]"
                message += "本日の順位計測データが存在しませんでした。[/info]"
                sendChatworkNotification(message)
                logger.debug(f'getUploadData: 本日の順位計測データが存在しませんでした。')
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
            
            message += f'『{keyword}』 {rank}位{medal}{arrow}%0a'

        message += "%0a＼check／✌🏻%0a"
        message += "https://aimplace.co.jp/p"
        return message
    except Exception as err:
        logger.debug(f'Error: getUploadData: {err}')
        exit(1)

def tweet_ranking_data(message):
    url = f'https://twitter.com/intent/tweet?text={message}'
    login = os.environ['TWITTER_PEOPLES_ID']
    password = os.environ['TWITTER_PEOPLES_PASS']

    try:
        driver = webdriver.Chrome(ChromeDriverManager().install())
        
        driver.get(url)
        driver.maximize_window()
        driver.implicitly_wait(5)

        driver.find_element_by_xpath('//input[@name="session[username_or_email]"]').send_keys(login)
        driver.implicitly_wait(5)
        driver.find_element_by_xpath('//input[@name="session[password]"]').send_keys(password)
        driver.implicitly_wait(5)
        driver.find_elements_by_xpath('//div[@role="button"]')[1].click()
        driver.implicitly_wait(10)
        
        driver.find_element_by_xpath('//div[@data-testid="scheduleOption"]').click()
        driver.implicitly_wait(5)

        dropdown = driver.find_element_by_id('SELECTOR_1') 
        select = Select(dropdown)
        select.select_by_value(today.strftime("%m"))
        driver.implicitly_wait(3)
        dropdown = driver.find_element_by_id('SELECTOR_2') 
        select = Select(dropdown)
        select.select_by_value(today.strftime("%d"))
        driver.implicitly_wait(3)
        dropdown = driver.find_element_by_id('SELECTOR_3') 
        select = Select(dropdown)
        select.select_by_value(today.strftime("%Y"))
        driver.implicitly_wait(3)
        dropdown = driver.find_element_by_id('SELECTOR_4') 
        select = Select(dropdown)
        select.select_by_value(today.strftime("12"))
        driver.implicitly_wait(3)
        dropdown = driver.find_element_by_id('SELECTOR_5') 
        select = Select(dropdown)
        select.select_by_value(today.strftime("0"))
        driver.implicitly_wait(3)

        driver.find_elements_by_xpath('//div[@role="button"]')[3].click()
        driver.implicitly_wait(10)
        driver.find_element_by_xpath('//div[@data-testid="tweetButton"]').click()
        sleep(5)

        message = "[info][title]PeoPle'sランキングツイート処理結果[/title]"
        message += "予約ツイートを12:00AMに設定しました。[/info]"
        sendChatworkNotification(message)
    except Exception as err:
        logger.debug(f'Error: tweet_ranking_data: {err}')
        exit(1)

### main_script ###
if __name__ == '__main__':

    try:
        rankDataDirPath = os.environ["RANK_DATA_DIR"]
        dateDirPath = getLatestDownloadedDirName(rankDataDirPath)

        data = list(getRankingCsvData(f'{dateDirPath}/aimplace.co.jp.txt'))
        logger.info(f'ranking: {data}')
        message = getUploadData(data)

        tweet_ranking_data(message)
        logger.info("tweet_ranking_data: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'tweet_ranking_data: {err}')
        exit(1)
