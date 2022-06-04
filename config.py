import sys
import warnings
warnings.simplefilter("ignore", UserWarning)
sys.coinit_flags = 2
import os
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import time
import pandas as pd
from PyQt5.QtGui import *
from pywinauto import application
from pywinauto import timings
from pywinauto.keyboard import send_keys
from datetime import datetime, timedelta
from openpyxl import load_workbook
from kiwoom import *
import xlrd
import shutil
import xlwings as xw

# 비밀 번호 정보
PASS = 'qwer1234'
PASSID = 'pq13!@zp7t'

# output specification
output_list = {
    '10062': ['종목코드',
                 '순위',
                 '종목명',
                 '현재가',
                 '대비기호',
                 '전일대비',
                 '등락률',
                 '누적거래량',
                 '기관순매매수량',
                 '기관순매매금액',
                 '기관순매매평균가',
                 '외인순매매수량',
                 '외인순매매금액',
                 '외인순매매평균가',
                 '순매매수량',
                 '순매매금액',                
                 ],
    '10066': ['종목코드',
                 '종목명',
                 '현재가',
                 '대비기호',
                 '전일대비',
                 '등락률',
                 '거래량',
                 '개인투자자',
                 '외국인투자자',
                 '기관계',
                 '금융투자',
                 '보험',
                 '투신',
                 '기타금융',
                 '은행',
                 '연기금등',
                 '사모펀드',
                 '국가',
                 '기타법인',               
                 ],
    '10059': [
                '현재가',
                 '대비기호',
                 '전일대비',
                 '등락률',
                 '누적거래량',
                 '누적거래대금',
                 '개인투자자',
                 '외국인투자자',
                 '기관계',
                 '금융투자',
                 '보험',
                 '투신',
                 '기타금융',
                 '은행',
                 '연기금등',
                 '사모펀드',
                 '국가',
                 '기타법인',
                 '내외국인',               
                 ],

    '10047': ['체결강도'
                ],
}