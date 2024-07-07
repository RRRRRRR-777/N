# moduleをインポート
import datetime
import glob
import math
import os
import re
import shutil
import sys
import time
from datetime import timedelta
from logging import getLogger, StreamHandler, Formatter
from statistics import mean
from dotenv import load_dotenv

import matplotlib.cm as cm
import numpy as np
import openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from matplotlib.colors import TwoSlopeNorm
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.borders import Border, Side
from selenium import webdriver
from selenium.webdriver.chrome import service as fs
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from utils import InitProcess


class PerfInst:
    def __init__(self, logger):
        self.logger = logger
        dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
        load_dotenv(dotenv_path)  # .envファイルから値を読み込む
        self.base_url = os.getenv("INSTITUTION_URL")
        self.search_url = f"{self.base_url}/search?q="

    # 検索画面から対象の銘柄のURLを検索する
    def search(self, ticker_code):
        search_url = self.search_url + ticker_code
        response = requests.get(search_url)
        soup = BeautifulSoup(response.text, "html.parser")

        a_tags = soup.find_all('li')
        for tag in a_tags:
            a_tag = tag.find('a')
            if ticker_code.upper() == a_tag.text.split()[0]:
                href = a_tag.get('href')
                if href:
                    full_url = f"{self.base_url}{href}"
                    response = requests.get(full_url)
                    if response.status_code == 200:
                        self.logger.info(f"Successfully accessed {full_url}")
                    else:
                        self.logger.error(f"Failed to access {full_url}")
                    return response
        self.logger.warning("No matching link found.")
        return None

    def getInst(self, ticker_code, response):
        soup = BeautifulSoup(response.text, "html.parser")
        tbody_tag = soup.find('tbody')
        previous_inst = tbody_tag.find_all('tr')[1].find_all('td')[1].text
        previous_inst_2 = tbody_tag.find_all('tr')[2].find_all('td')[1].text
        print(f"前期: {previous_inst}, 前々期: {previous_inst_2}")

    def execute(self):
        ticker_list = ["aapl", "meta", "amzn", "nvda", "tsla", "app", "tko"]
        for ticker_code in ticker_list:
            response = self.search(ticker_code)
            if response:
                self.getInst(ticker_code, response)
            time.sleep(1.0)


# インスタンスの作成と実行
init_process = InitProcess()
logger = init_process.set_log()
perf_inst = PerfInst(logger)
perf_inst.execute()
