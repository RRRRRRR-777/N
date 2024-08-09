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
                    if not response.status_code == 200:
                        self.logger.error(f"Failed to access {full_url}")
                    return response
        self.logger.warning("No matching link found.")
        return None

    # 対象銘柄の機関投資家購入数の前期と前々期を取得
    def getInst(self, ticker_code, response):
        try:
            soup = BeautifulSoup(response.text, "html.parser")
            tbody_tag = soup.find('tbody')
            previous_inst_ = tbody_tag.find_all('tr')[1].find_all('td')[1].text
            previous_inst = int(previous_inst_.replace("\n", "").replace(",", ""))
            previous_inst_2_ = tbody_tag.find_all('tr')[2].find_all('td')[1].text
            previous_inst_2 = int(previous_inst_2_.replace("\n", "").replace(",", ""))
            return previous_inst, previous_inst_2
        except Exception as e:
            self.logger.warning(f"No:{i+1}, Ticker:{ticker_code}, {e}")
            return None, None

    def execute(self):
        # 銘柄のリストを読み込む
        stock_df = pd.read_csv(os.path.join(
            glob.glob(os.getcwd()+"/CsvData*")[0], "BuyingStock.csv"))
        stock_df_ticker = stock_df["Ticker"]
        # 取得したデータを格納する配列
        data = []
        for i, ticker_code in enumerate(stock_df_ticker):
            # 検索画面から対象の銘柄のURLを検索する
            response = self.search(ticker_code)
            if response:
                # 対象銘柄の機関投資家購入数の前期と前々期を取得
                previous_inst, previous_inst_2 = self.getInst(ticker_code, response)
                data.append(
                    [ticker_code, previous_inst, previous_inst_2])
            time.sleep(1.0)
            print(f"Done {i+1}/{len(stock_df_ticker)}", end="\r")
        # dataをdfに変換する
        df = pd.DataFrame(data)
        col = ["Ticker", "PreviousInst", "PreviousInst2"]
        df.columns = col
        # 前期と前々期のパフォーマンスを計算
        performance = (df['PreviousInst'] -
                       df['PreviousInst2']) / abs(df['PreviousInst2'])
        df["PerfInst"] = round(performance * 100, 2)
        # dfをCSV出力する
        outputDir = os.path.join(
            glob.glob(os.getcwd()+"/CsvData*")[0], f"PerfInst.csv")
        df.to_csv(outputDir, index=False)

# インスタンスの作成と実行
init_process = InitProcess()
logger = init_process.set_log()
perf_inst = PerfInst(logger)
perf_inst.execute()
