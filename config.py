"""
配置檔案 - 集中管理所有常數和設定
"""
import os
import sys

###########################################################################
# 路徑設定（支援 exe 打包）
###########################################################################

def get_base_dir():
    """取得程式執行的基礎目錄"""
    if getattr(sys, 'frozen', False):
        # 如果是打包後的 exe
        return os.path.dirname(sys.executable)
    else:
        # 如果是 Python 腳本
        return os.path.dirname(os.path.abspath(__file__))

# 全域路徑常數
BASE_DIR = get_base_dir()
DATA_DIR = os.path.join(BASE_DIR, 'data')
LOGS_DIR = os.path.join(BASE_DIR, 'logs')

# 快取目錄
REVENUE_CACHE_DIR = os.path.join(DATA_DIR, 'revenue')
FINANCIAL_CACHE_DIR = os.path.join(DATA_DIR, 'financial')

###########################################################################
# API 設定
###########################################################################

# FinMind API Token（如需要可在此設定）
API_TOKEN = ""

###########################################################################
# 日誌設定
###########################################################################

# 日誌保留天數
LOG_RETENTION_DAYS = 7

# 日誌檔案名稱格式
LOG_FILENAME_FORMAT = 'stock_processor_%Y%m%d.log'

