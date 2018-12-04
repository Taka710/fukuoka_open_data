import os
from glob import glob
from datetime import datetime
from time import sleep
import requests

import xlrd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from bs4 import BeautifulSoup

from fukuoka_populatiton.const import *


def get_popular_file_path():
    """
    excelファイルのファイルパスリストを取得
    :return: xls file Path
    """
    main_path = os.path.dirname(os.path.abspath(__name__))
    xlsx_path = os.path.join(main_path, 'xlsx/*')

    for xls_list in glob(xlsx_path):
        yield xls_list


def get_popular_data(path: str):
    """
    xlsファイルから必要なデータを取得
    :param path: xlsファイルパス
    :return: xls行データ
    """
    str_file_date = str(os.path.basename(path)).split('-')

    file_date = datetime(int(str_file_date[2]),
                         int(str_file_date[1]),
                         int(str_file_date[0]))

    xls_book = xlrd.open_workbook(path)
    sheets = xls_book.sheets()
    sheet = sheets[0]

    for i in range(EXCEL_START_ROW, sheet.nrows - 1):

        # データを除外する一部の市、郡をスキップ
        if sheet.cell_value(i, 2) in DISTRICT:
            continue
        else:
            xls_data = [
                file_date.strftime('%Y-%m-%d'),
                sheet.cell_value(i, 2),
                sheet.cell_value(i, 6),
                sheet.cell_value(i, 7),
                sheet.cell_value(i, 8),
                sheet.cell_value(i, 9)
            ]

        yield xls_data


def get_area_data(*args):
    """
    地域の面積データを取得
    :param args: 未使用（関数の書式を共通化するために定義）
    :return: 地域名と面積データ
    """

    url = AREA_URL
    res = requests.get(url)
    soup = BeautifulSoup(res.content, 'html.parser')

    table = soup.findAll("table", {"class": "t51"})[0]
    rows = table.findAll("tr")

    for i, area_row in enumerate(rows):

        if i == 0:
            continue

        town_data = []

        for cell in area_row.findAll("td", {"class": "al bcity pd13 no"}):
            town_data.append(cell.string)

        for cell in area_row.findAll("td", {"class": "al bseku pd13 no"}):
            town_data.append(cell.string)

        for cell in area_row.findAll("td", {"class": "al btown pd13 no"}):
            town_data.append(cell.string)

        for cell in area_row.findAll("td", {"class": "al bvill pd13 no"}):
            town_data.append(cell.string)

        for cell in area_row.findAll("td", {"class": "ar bw pd13 no"}):
            town_data.append(cell.string)

        yield town_data


def set_gspread(path: str, start_row: int, sheet_name: str, func):
    """
    google spreadにデータを書き込み

    :param path: データ取得元ファイルpath
    :param start_row: spread書き込み開始行
    :param sheet_name: spred書き込みsheet名
    :param func: データ取得関数
    :return: 次回書き込み開始行
    """

    # 初期設定
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    main_path = os.path.dirname(os.path.abspath(__name__))
    json_path = os.path.join(main_path, 'json/*')

    json_files = glob(json_path)

    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_files[0], scope)
    gc = gspread.authorize(credentials)
    google_sheet = gc.open(GOOGLE_SPREAD_NAME).worksheet(sheet_name)

    next_row = start_row

    for index, record in enumerate(func(path)):
        # 更新対象列を取得
        cell_list = google_sheet.range(index + start_row, 1, index + start_row, len(record))

        # セルの更新を行う
        for cell, val in zip(cell_list, record):
            cell.value = val

        # セルの更新
        google_sheet.update_cells(cell_list)

        next_row += 1

    return next_row


if __name__ == '__main__':

    # 人口データを登録
    row = 2
    for file_path in get_popular_file_path():
        row = set_gspread(file_path, row, POPULAR_SPREAD, get_popular_data)
        sleep(60)

    # 面積データを登録
    row = 1
    set_gspread("", row, AREA_SPREAD, get_area_data)
