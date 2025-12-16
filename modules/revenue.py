"""
營收數據處理模組
"""
import logging
from datetime import datetime

from modules.cache import has_latest_revenue, load_cache, save_cache


def get_stock_revenue_data(api, stock_id, start_date=None, use_cache=True):
    """獲取股票營收數據（帶快取）"""
    # 檢查地端是否有上個月營收
    if use_cache and has_latest_revenue(stock_id):
        cached_data = load_cache(stock_id, 'revenue')
        if cached_data is not None:
            logging.info(f"  ✓ 快取: {stock_id} 營收")
            return cached_data
    
    # 地端沒有上個月資料，從 API 抓取
    if start_date is None:
        two_years_ago = datetime.now().replace(year=datetime.now().year - 2)
        start_date = two_years_ago.strftime('%Y-%m-%d')
    
    logging.info(f"  ⟳ API: {stock_id} 營收")
    data = api.taiwan_stock_month_revenue(
        stock_id=stock_id,
        start_date=start_date,
    )
    
    # 儲存到快取
    if data is not None and not data.empty:
        save_cache(stock_id, 'revenue', data)
    
    return data


def extract_revenue_by_year_month(revenue_data, target_year, target_month):
    """從營收數據中提取指定年月的營收"""
    for i in range(len(revenue_data) - 1, -1, -1):
        revenue_year = int(revenue_data.iloc[i]['revenue_year'])
        revenue_month = int(revenue_data.iloc[i]['revenue_month'])
        if revenue_year == target_year and revenue_month == target_month:
            return revenue_data.iloc[i]['revenue']
    return None


def get_previous_two_months():
    """取得上個月和上上個月的年份和月份（保留舊函數以維持相容性）"""
    result = get_previous_three_months()
    return result[0], result[1]


def get_previous_three_months():
    """取得上個月、上上個月和上上上個月的年份和月份"""
    current_year = datetime.now().year
    current_month = datetime.now().month
    
    # 計算上個月
    if current_month > 1:
        last_month = current_month - 1
        last_month_year = current_year
    else:
        last_month = 12
        last_month_year = current_year - 1
    
    # 計算上上個月
    if last_month > 1:
        previous_month = last_month - 1
        previous_month_year = last_month_year
    else:
        previous_month = 12
        previous_month_year = last_month_year - 1
    
    # 計算上上上個月
    if previous_month > 1:
        previous_month2 = previous_month - 1
        previous_month_year2 = previous_month_year
    else:
        previous_month2 = 12
        previous_month_year2 = previous_month_year - 1
    
    return (last_month_year, last_month), (previous_month_year, previous_month), (previous_month_year2, previous_month2)


def get_ytd_revenue_from_monthly(revenue_data):
    """從月營收數據計算今年累積營收"""
    current_year = datetime.now().year
    ytd_revenue_data = revenue_data[revenue_data['revenue_year'] == current_year]
    
    if ytd_revenue_data.empty:
        return None
    
    return ytd_revenue_data['revenue'].sum()


def get_ytd_revenue_yoy(revenue_data):
    """計算今年累積營收YoY，根據實際有資料的月份進行比較"""
    current_year = datetime.now().year
    last_year = current_year - 1
    
    # 取得今年的營收數據
    current_year_data = revenue_data[revenue_data['revenue_year'] == current_year]
    
    if current_year_data.empty:
        return None, None
    
    # 找出今年最新有資料的月份
    latest_month = current_year_data['revenue_month'].max()
    
    # 計算今年截至最新月份的累積營收
    current_ytd_data = current_year_data[current_year_data['revenue_month'] <= latest_month]
    current_ytd = current_ytd_data['revenue'].sum()
    
    # 計算去年同期（相同月份）的累積營收
    last_year_data = revenue_data[
        (revenue_data['revenue_year'] == last_year) &
        (revenue_data['revenue_month'] <= latest_month)
    ]
    
    if last_year_data.empty:
        return None, latest_month
    
    last_year_ytd = last_year_data['revenue'].sum()
    
    # 計算YoY
    if last_year_ytd and last_year_ytd != 0:
        yoy = round((current_ytd - last_year_ytd) / last_year_ytd * 100, 2)
        return yoy, latest_month
    
    return None, latest_month


def process_revenue_data(api, df, idx, stock_id, last_month_year, last_month, previous_month_year, previous_month, previous_month_year2, previous_month2, yoy_year):
    """處理單一股票的營收數據"""
    from modules.utils import convert_to_million, ensure_column_exists
    
    # 動態初始化所有營收相關欄位（與 financial/eps 模組保持一致）
    current_year = datetime.now().year
    ensure_column_exists(df, f'{last_month}月營收(M)')
    ensure_column_exists(df, f'{previous_month}月營收(M)')
    ensure_column_exists(df, f'{previous_month2}月營收(M)')
    ensure_column_exists(df, 'MoM(%)')
    ensure_column_exists(df, 'YoY(%)')
    ensure_column_exists(df, f'{str(current_year)[-2:]}年累積營收(M)')
    ensure_column_exists(df, '累積營收YoY(%)')
    
    try:
        revenue_data = get_stock_revenue_data(api, stock_id)
        
        if revenue_data is None or revenue_data.empty:
            logging.warning(f"  警告: {stock_id} 無營收數據")
            return
    except Exception as e:
        logging.error(f"  錯誤: {stock_id} 營收數據獲取失敗 - {str(e)}")
        return
    
    revenue_current = extract_revenue_by_year_month(revenue_data, last_month_year, last_month)
    revenue_previous = extract_revenue_by_year_month(revenue_data, previous_month_year, previous_month)
    revenue_previous2 = extract_revenue_by_year_month(revenue_data, previous_month_year2, previous_month2)
    revenue_yoy = extract_revenue_by_year_month(revenue_data, yoy_year, last_month)
    
    # 轉換為百萬單位
    revenue_current_million = convert_to_million(revenue_current)
    revenue_previous_million = convert_to_million(revenue_previous)
    revenue_previous2_million = convert_to_million(revenue_previous2)
    
    # 計算 MoM (Month over Month)
    if revenue_current and revenue_previous and revenue_previous != 0:
        mom = round((revenue_current - revenue_previous) / revenue_previous * 100, 2)
    else:
        mom = None
    
    # 計算 YoY (Year over Year)
    if revenue_current and revenue_yoy and revenue_yoy != 0:
        yoy = round((revenue_current - revenue_yoy) / revenue_yoy * 100, 2)
    else:
        yoy = None
    
    # 計算今年累積營收
    ytd_revenue = get_ytd_revenue_from_monthly(revenue_data)
    ytd_revenue_million = convert_to_million(ytd_revenue)
    
    # 計算累積營收YoY
    ytd_yoy, ytd_month = get_ytd_revenue_yoy(revenue_data)
    
    # 更新 DataFrame
    current_year = datetime.now().year
    df.at[idx, f'{last_month}月營收(M)'] = revenue_current_million
    df.at[idx, f'{previous_month}月營收(M)'] = revenue_previous_million
    df.at[idx, f'{previous_month2}月營收(M)'] = revenue_previous2_million
    df.at[idx, 'MoM(%)'] = mom
    df.at[idx, 'YoY(%)'] = yoy
    df.at[idx, f'{str(current_year)[-2:]}年累積營收(M)'] = ytd_revenue_million
    df.at[idx, '累積營收YoY(%)'] = ytd_yoy
