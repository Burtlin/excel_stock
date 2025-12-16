"""
財務報表數據處理模組
"""
import logging
from datetime import datetime

from modules.cache import has_latest_financial, load_cache, save_cache


def get_last_season_month():
    """取得上一季的季末月份和年份"""
    current_month = datetime.now().month
    season_months = [3, 6, 9, 12]
    
    # 找出上一季的季末月份
    last_season_month = None
    for month in reversed(season_months):
        if current_month > month:
            last_season_month = month
            break
    
    # 如果當前月份小於等於3月，上一季是去年12月
    if last_season_month is None:
        last_season_month = 12
        target_year = datetime.now().year - 1
    else:
        target_year = datetime.now().year
    
    return target_year, last_season_month


def get_stock_financial_data(api, stock_id, start_date=None, use_cache=True):
    """獲取股票財務報表數據（帶快取）"""
    # 檢查地端是否有上季財務資料
    if use_cache and has_latest_financial(stock_id):
        cached_data = load_cache(stock_id, 'financial')
        if cached_data is not None:
            logging.info(f"  ✓ 快取: {stock_id} 財務")
            return cached_data
    
    # 地端沒有上季資料，從 API 抓取
    if start_date is None:
        two_years_ago = datetime.now().replace(year=datetime.now().year - 2)
        start_date = two_years_ago.strftime('%Y-%m-%d')
    
    logging.info(f"  ⟳ API: {stock_id} 財務")
    data = api.taiwan_stock_financial_statement(
        stock_id=stock_id,
        start_date=start_date,
    )
    
    # 儲存到快取
    if data is not None and not data.empty:
        save_cache(stock_id, 'financial', data)
    
    return data


def extract_value_by_date(financial_data, data_type, target_date):
    """從財務數據中提取指定類型和日期的數據（優化版）"""
    filtered_data = financial_data[
        (financial_data['type'] == data_type) & 
        (financial_data['date'] == target_date)
    ]
    return filtered_data.iloc[0]['value'] if not filtered_data.empty else None


def get_quarter_name(target_year, season_month):
    """根據年份和月份取得季度名稱"""
    quarter_map = {3: 'Q1', 6: 'Q2', 9: 'Q3', 12: 'Q4'}
    return f"{str(target_year)[-2:]}{quarter_map[season_month]}"


def get_season_date(target_year, target_month):
    """組合季度結束日期"""
    if target_month in [6, 9]:
        return f"{target_year}-{target_month:02d}-30"
    else:  # 3月、12月
        return f"{target_year}-{target_month:02d}-31"


def get_previous_season_month(target_year, target_month):
    """取得上上季的年份和月份"""
    season_months = [3, 6, 9, 12]
    current_index = season_months.index(target_month)
    
    if current_index > 0:
        return target_year, season_months[current_index - 1]
    else:
        return target_year - 1, 12


def get_last_two_season_data(financial_data, data_type):
    """通用函數：取得上一季和上上季的指定數據"""
    # 計算上一季和上上季的年份月份
    target_year, last_season_month = get_last_season_month()
    prev_year, prev_month = get_previous_season_month(target_year, last_season_month)
    
    # 組合目標日期
    last_date = get_season_date(target_year, last_season_month)
    prev_date = get_season_date(prev_year, prev_month)
    
    # 提取兩季的數據
    value_last = extract_value_by_date(financial_data, data_type, last_date)
    value_prev = extract_value_by_date(financial_data, data_type, prev_date)
    
    # 取得季度名稱
    quarter_last = get_quarter_name(target_year, last_season_month)
    quarter_prev = get_quarter_name(prev_year, prev_month)
    
    return value_last, quarter_last, value_prev, quarter_prev


def get_last_three_season_data(financial_data, data_type):
    """通用函數：取得上一季、上上季和上上上季的指定數據"""
    # 計算上一季、上上季和上上上季的年份月份
    target_year, last_season_month = get_last_season_month()
    prev_year, prev_month = get_previous_season_month(target_year, last_season_month)
    prev2_year, prev2_month = get_previous_season_month(prev_year, prev_month)
    
    # 組合目標日期
    last_date = get_season_date(target_year, last_season_month)
    prev_date = get_season_date(prev_year, prev_month)
    prev2_date = get_season_date(prev2_year, prev2_month)
    
    # 提取三季的數據
    value_last = extract_value_by_date(financial_data, data_type, last_date)
    value_prev = extract_value_by_date(financial_data, data_type, prev_date)
    value_prev2 = extract_value_by_date(financial_data, data_type, prev2_date)
    
    # 取得季度名稱
    quarter_last = get_quarter_name(target_year, last_season_month)
    quarter_prev = get_quarter_name(prev_year, prev_month)
    quarter_prev2 = get_quarter_name(prev2_year, prev2_month)
    
    return value_last, quarter_last, value_prev, quarter_prev, value_prev2, quarter_prev2


def calculate_gross_margin(financial_data):
    """計算所有季度的毛利率"""
    # 提取毛利和營收數據
    gross_profit_data = financial_data[financial_data['type'] == 'GrossProfit'][['date', 'value']].rename(columns={'value': 'gross_profit'})
    revenue_data = financial_data[financial_data['type'] == 'Revenue'][['date', 'value']].rename(columns={'value': 'revenue'})
    
    # 合併數據
    merged_data = gross_profit_data.merge(revenue_data, on='date', how='inner')
    
    # 計算毛利率
    merged_data['gross_margin'] = (merged_data['gross_profit'] / merged_data['revenue'] * 100).round(2)
    
    return merged_data


def get_ytd_revenue(financial_data):
    """計算今年累積營收（Year-To-Date Revenue）"""
    current_year = datetime.now().year
    
    # 篩選今年的 Revenue 數據
    revenue_data = financial_data[
        (financial_data['type'] == 'Revenue') &
        (financial_data['date'].str.startswith(str(current_year)))
    ]
    
    if revenue_data.empty:
        return None
    
    # 加總今年所有季度的營收
    ytd_revenue = revenue_data['value'].sum()
    return ytd_revenue


def get_last_two_season_gross_margin(financial_data):
    """取得股票上一季和上上季的毛利率"""
    # 計算上一季和上上季的年份月份
    target_year, last_season_month = get_last_season_month()
    prev_year, prev_month = get_previous_season_month(target_year, last_season_month)
    
    # 組合目標日期
    last_date = get_season_date(target_year, last_season_month)
    prev_date = get_season_date(prev_year, prev_month)
    
    # 計算毛利率
    gross_margin_data = calculate_gross_margin(financial_data)
    
    # 提取指定日期的毛利率
    gross_margin_last = gross_margin_data[gross_margin_data['date'] == last_date]['gross_margin'].values
    gross_margin_prev = gross_margin_data[gross_margin_data['date'] == prev_date]['gross_margin'].values
    
    gross_margin_last = gross_margin_last[0] if len(gross_margin_last) > 0 else None
    gross_margin_prev = gross_margin_prev[0] if len(gross_margin_prev) > 0 else None
    
    # 取得季度名稱
    quarter_last = get_quarter_name(target_year, last_season_month)
    quarter_prev = get_quarter_name(prev_year, prev_month)
    
    return gross_margin_last, quarter_last, gross_margin_prev, quarter_prev


def get_ytd_eps(financial_data):
    """計算今年累積EPS（Year-To-Date EPS）"""
    current_year = datetime.now().year
    
    # 篩選今年的 EPS 數據
    eps_data = financial_data[
        (financial_data['type'] == 'EPS') &
        (financial_data['date'].str.startswith(str(current_year)))
    ]
    
    if eps_data.empty:
        return None
    
    # 加總今年所有季度的 EPS
    ytd_eps = eps_data['value'].sum()
    return round(ytd_eps, 2) if ytd_eps else None


def process_financial_data(api, df, idx, stock_id):
    """處理單一股票的綜合損益表（季營收、毛利率、累積營收）"""
    from datetime import datetime
    from modules.utils import convert_to_million, ensure_column_exists
    
    financial_data = get_stock_financial_data(api, stock_id)
    if financial_data is None or financial_data.empty:
        logging.warning(f"  警告: {stock_id} 無財務數據")
        return
    
    # 處理季營收（從財務報表的 Revenue 提取）
    season_revenue_last, sr_quarter_last, season_revenue_prev, sr_quarter_prev = get_last_two_season_data(financial_data, 'Revenue')
    
    # 轉換為百萬單位
    season_revenue_last_million = convert_to_million(season_revenue_last)
    season_revenue_prev_million = convert_to_million(season_revenue_prev)
    
    # 初始化並更新季營收欄位
    ensure_column_exists(df, f'{sr_quarter_last}季營收(M)')
    ensure_column_exists(df, f'{sr_quarter_prev}季營收(M)')
    df.at[idx, f'{sr_quarter_last}季營收(M)'] = season_revenue_last_million
    df.at[idx, f'{sr_quarter_prev}季營收(M)'] = season_revenue_prev_million
    
    # 處理毛利率
    gross_margin_last, gm_quarter_last, gross_margin_prev, gm_quarter_prev = get_last_two_season_gross_margin(financial_data)
    
    # 初始化並更新毛利率欄位
    ensure_column_exists(df, f'{gm_quarter_last}毛利率(%)')
    ensure_column_exists(df, f'{gm_quarter_prev}毛利率(%)')
    df.at[idx, f'{gm_quarter_last}毛利率(%)'] = gross_margin_last
    df.at[idx, f'{gm_quarter_prev}毛利率(%)'] = gross_margin_prev
    
    # 處理今年累積營收
    ytd_revenue = get_ytd_revenue(financial_data)
    ytd_revenue_million = convert_to_million(ytd_revenue)
    
    # 初始化並更新累積營收欄位
    current_year = datetime.now().year
    ensure_column_exists(df, f'{str(current_year)[-2:]}年累積營收(M)')
    df.at[idx, f'{str(current_year)[-2:]}年累積營收(M)'] = ytd_revenue_million


def process_eps_data(api, df, idx, stock_id):
    """處理單一股票的 EPS 數據"""
    from datetime import datetime
    from modules.utils import ensure_column_exists
    
    financial_data = get_stock_financial_data(api, stock_id)
    if financial_data is None or financial_data.empty:
        logging.warning(f"  警告: {stock_id} 無財務數據")
        return
    
    # 處理 EPS（取得三季數據）
    try:
        eps_last, quarter_last, eps_prev, quarter_prev, eps_prev2, quarter_prev2 = get_last_three_season_data(financial_data, 'EPS')
    except Exception as e:
        logging.error(f"  錯誤: {stock_id} EPS 提取失敗 - {str(e)}")
        eps_last = quarter_last = eps_prev = quarter_prev = eps_prev2 = quarter_prev2 = None
    
    # 計算今年累積 EPS
    try:
        ytd_eps = get_ytd_eps(financial_data)
    except Exception as e:
        logging.error(f"  錯誤: {stock_id} YTD EPS 計算失敗 - {str(e)}")
        ytd_eps = None
    
    # 初始化並更新 EPS 欄位
    ensure_column_exists(df, f'{quarter_last}EPS')
    ensure_column_exists(df, f'{quarter_prev}EPS')
    ensure_column_exists(df, f'{quarter_prev2}EPS')
    current_year = datetime.now().year
    ensure_column_exists(df, f'{str(current_year)[-2:]}年累積EPS')
    df.at[idx, f'{quarter_last}EPS'] = eps_last
    df.at[idx, f'{quarter_prev}EPS'] = eps_prev
    df.at[idx, f'{quarter_prev2}EPS'] = eps_prev2
    df.at[idx, f'{str(current_year)[-2:]}年累積EPS'] = ytd_eps
