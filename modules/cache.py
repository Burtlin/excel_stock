"""
快取管理模組
"""
import json
import logging
import os
from datetime import datetime

import pandas as pd

from config import DATA_DIR


def get_cache_path(stock_id, data_type):
    """取得快取檔案路徑"""
    cache_dir = os.path.join(DATA_DIR, data_type)
    os.makedirs(cache_dir, exist_ok=True)
    return os.path.join(cache_dir, f'{stock_id}.json')


def has_latest_revenue(stock_id):
    """檢查快取中是否有上個月的營收資料"""
    cache_path = get_cache_path(stock_id, 'revenue')
    if not os.path.exists(cache_path):
        return False
    
    try:
        with open(cache_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        df = pd.DataFrame(data)
        
        if df.empty:
            return False
        
        # 計算上個月
        current_year = datetime.now().year
        current_month = datetime.now().month
        if current_month > 1:
            last_month = current_month - 1
            last_month_year = current_year
        else:
            last_month = 12
            last_month_year = current_year - 1
        
        # 檢查是否有上個月的資料
        has_data = any(
            (int(row['revenue_year']) == last_month_year and 
             int(row['revenue_month']) == last_month)
            for _, row in df.iterrows()
        )
        return has_data
    except:
        return False


def has_latest_financial(stock_id):
    """檢查快取中是否有上季的財務資料"""
    cache_path = get_cache_path(stock_id, 'financial')
    if not os.path.exists(cache_path):
        return False
    
    try:
        with open(cache_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        df = pd.DataFrame(data)
        
        if df.empty:
            return False
        
        # 計算上一季
        current_month = datetime.now().month
        season_months = [3, 6, 9, 12]
        last_season_month = None
        for month in reversed(season_months):
            if current_month > month:
                last_season_month = month
                break
        
        if last_season_month is None:
            last_season_month = 12
            target_year = datetime.now().year - 1
        else:
            target_year = datetime.now().year
        
        # 組合目標日期
        if last_season_month in [6, 9]:
            target_date = f"{target_year}-{last_season_month:02d}-30"
        else:
            target_date = f"{target_year}-{last_season_month:02d}-31"
        
        # 檢查是否有上一季的資料
        has_data = any(
            row['date'] == target_date
            for _, row in df.iterrows()
        )
        return has_data
    except:
        return False


def load_cache(stock_id, data_type):
    """從快取載入數據"""
    cache_path = get_cache_path(stock_id, data_type)
    
    if os.path.exists(cache_path):
        with open(cache_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return pd.DataFrame(data)
    return None


def save_cache(stock_id, data_type, df):
    """儲存數據到快取"""
    cache_path = get_cache_path(stock_id, data_type)
    df.to_json(cache_path, orient='records', force_ascii=False, indent=2)
