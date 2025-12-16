"""
通用工具函數模組
"""
import json
import os

import pandas as pd

from config import DATA_DIR


def get_stock_name_mapping(api):
    """獲取股票代號與名稱的對應字典"""
    stock_info_path = os.path.join(DATA_DIR, 'stock_info.json')
    if os.path.exists(stock_info_path):
        with open(stock_info_path, 'r', encoding='utf-8') as f:
            stock_dict = json.load(f)
        return stock_dict
    
    df = api.taiwan_stock_info()
    stock_dict = dict(zip(df['stock_id'], df['stock_name']))
    # save to json
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(stock_info_path, 'w', encoding='utf-8') as f:
        json.dump(stock_dict, f, ensure_ascii=False, indent=4)
    return stock_dict


def add_stock_names(df, stock_dict):
    """根據股票代號加入名稱欄位"""
    df['名稱'] = df['代號'].astype(str).map(stock_dict)
    return df


def process_info_data(api, df):
    """處理股票資訊數據，加入名稱欄位"""
    stock_dict = get_stock_name_mapping(api)
    df = add_stock_names(df, stock_dict)
    return df


def convert_to_million(value):
    """將數值從元轉換為百萬單位（取整）"""
    return round(value / 1000000) if value else None


def ensure_column_exists(df, column_name):
    """確保欄位存在，如果不存在則初始化"""
    if column_name not in df.columns:
        df[column_name] = None


def format_percentage_columns(worksheet, df):
    """格式化工作表中的百分比欄位"""
    from openpyxl.utils import get_column_letter
    
    for col_idx, col_name in enumerate(df.columns, start=1):
        if '(%)' in str(col_name):
            col_letter = get_column_letter(col_idx)
            for row in range(2, len(df) + 2):
                cell = worksheet[f'{col_letter}{row}']
                if cell.value is not None and isinstance(cell.value, (int, float)):
                    cell.number_format = '0.00%'
                    cell.value = cell.value / 100
