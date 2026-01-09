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


def get_ytd_revenue_from_monthly(revenue_data, target_year, target_month):
    """從月營收數據計算指定年份到指定月份的累積營收
    
    Args:
        revenue_data: 營收數據
        target_year: 目標年份
        target_month: 目標月份（計算到此月為止）
    """
    ytd_revenue_data = revenue_data[
        (revenue_data['revenue_year'] == target_year) &
        (revenue_data['revenue_month'] <= target_month)
    ]
    
    if ytd_revenue_data.empty:
        return None
    
    return ytd_revenue_data['revenue'].sum()


def get_ytd_revenue_yoy(revenue_data, target_year, target_month):
    """計算指定年份到指定月份的累積營收YoY
    
    Args:
        revenue_data: 營收數據
        target_year: 目標年份
        target_month: 目標月份（計算到此月為止）
    """
    last_year = target_year - 1
    
    # 計算目標年份截至目標月份的累積營收
    current_ytd_data = revenue_data[
        (revenue_data['revenue_year'] == target_year) &
        (revenue_data['revenue_month'] <= target_month)
    ]
    
    if current_ytd_data.empty:
        return None
    
    current_ytd = current_ytd_data['revenue'].sum()
    
    # 計算去年同期（相同月份）的累積營收
    last_year_data = revenue_data[
        (revenue_data['revenue_year'] == last_year) &
        (revenue_data['revenue_month'] <= target_month)
    ]
    
    if last_year_data.empty:
        return None
    
    last_year_ytd = last_year_data['revenue'].sum()
    
    # 計算YoY
    if last_year_ytd and last_year_ytd != 0:
        yoy = round((current_ytd - last_year_ytd) / last_year_ytd * 100, 2)
        return yoy
    
    return None


def process_revenue_data(api, df, idx, stock_id, last_month_year, last_month, previous_month_year, previous_month, previous_month_year2, previous_month2, yoy_year):
    """處理單一股票的營收數據"""
    from modules.utils import convert_to_million, ensure_column_exists
    
    # 動態初始化所有營收相關欄位（與 financial/eps 模組保持一致）
    ensure_column_exists(df, f'{last_month}月營收(M)')
    ensure_column_exists(df, f'{previous_month}月營收(M)')
    ensure_column_exists(df, f'{previous_month2}月營收(M)')
    ensure_column_exists(df, 'MoM(%)')
    ensure_column_exists(df, 'YoY(%)')
    ensure_column_exists(df, f'{str(last_month_year)[-2:]}年累積營收(M)')
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
    
    # 計算到上個月為止的累積營收
    ytd_revenue = get_ytd_revenue_from_monthly(revenue_data, last_month_year, last_month)
    ytd_revenue_million = convert_to_million(ytd_revenue)
    
    # 計算累積營收YoY
    ytd_yoy = get_ytd_revenue_yoy(revenue_data, last_month_year, last_month)
    
    # 更新 DataFrame
    df.at[idx, f'{last_month}月營收(M)'] = revenue_current_million
    df.at[idx, f'{previous_month}月營收(M)'] = revenue_previous_million
    df.at[idx, f'{previous_month2}月營收(M)'] = revenue_previous2_million
    df.at[idx, 'MoM(%)'] = mom
    df.at[idx, 'YoY(%)'] = yoy
    df.at[idx, f'{str(last_month_year)[-2:]}年累積營收(M)'] = ytd_revenue_million
    df.at[idx, '累積營收YoY(%)'] = ytd_yoy


def create_revenue_overview(api, df_base):
    """建立營收總覽（橫向格式：月份為欄，股票為列，每支股票3行）
    
    格式：
    代號 | 名稱   | 12月    | 11月    | 10月    | ... | 1月
    2330 | 台積電 | 343614  | 367473  | ...     | ... | ...
    MoM(%)|       | -6.51%  | 2.20%   | ...     | ... | ...
    YoY(%)|       | 14.12%  | 31.77%  | ...     | ... | ...
    2382 | 廣達   | 192947  | 173196  | ...     | ... | ...
    MoM(%)|       | 11.41%  | ...     | ...     | ... | ...
    YoY(%)|       | 16.23%  | ...     | ...     | ... | ...
    ...
    """
    import pandas as pd
    from modules.utils import get_stock_name_mapping
    
    # 取得股票名稱對應
    stock_dict = get_stock_name_mapping(api)
    
    # 取得最近12個月的月份列表（從最新月份往前推12個月）
    current_year = datetime.now().year
    current_month = datetime.now().month
    
    # 計算12個月份（由新到舊）
    months_list = []
    for i in range(12):
        year = current_year
        month = current_month - i
        
        if month <= 0:
            month += 12
            year -= 1
        
        months_list.append((year, month))
    
    # 建立欄位名稱（25年1月、25年11月、...、24年1月）
    columns = ['代號', '名稱'] + [f"{str(year)[-2:]}年{month}月" for year, month in months_list]
    
    # 處理每支股票（每支股票3行：營收、MoM、YoY）
    rows = []
    total = len(df_base)
    
    for idx, row in df_base.iterrows():
        stock_id = str(row["代號"])
        stock_name = stock_dict.get(stock_id, "未知")
        
        logging.info(f"  [{idx+1}/{total}] 處理營收總覽: {stock_id} {stock_name}")
        
        try:
            # 取得營收數據
            revenue_data = get_stock_revenue_data(api, stock_id)
            
            if revenue_data is None or revenue_data.empty:
                logging.warning(f"    警告: {stock_id} 無營收數據")
                # 建立空白的三行
                revenue_row = {'代號': stock_id, '名稱': stock_name}
                mom_row = {'代號': '', '名稱': 'MoM(%)'}
                yoy_row = {'代號': '', '名稱': 'YoY(%)'}
                
                for year, month in months_list:
                    month_label = f"{str(year)[-2:]}年{month}月"
                    revenue_row[month_label] = None
                    mom_row[month_label] = None
                    yoy_row[month_label] = None
                
                rows.extend([revenue_row, mom_row, yoy_row])
                continue
            
            # 第1行：營收數據
            revenue_row = {'代號': stock_id, '名稱': stock_name}
            # 第2行：MoM
            mom_row = {'代號': '', '名稱': 'MoM(%)'}
            # 第3行：YoY
            yoy_row = {'代號': '', '名稱': 'YoY(%)'}
            
            # 提取每個月份的營收並計算MoM和YoY
            for i, (year, month) in enumerate(months_list):
                month_label = f"{str(year)[-2:]}年{month}月"
                revenue = extract_revenue_by_year_month(revenue_data, year, month)
                
                # 營收（轉換為百萬單位）
                if revenue is not None:
                    revenue_million = round(revenue / 1000000)
                    revenue_row[month_label] = revenue_million
                    
                    # 計算 MoM（與上一個月比較）
                    # 計算上個月的年份和月份
                    prev_month = month - 1 if month > 1 else 12
                    prev_year = year if month > 1 else year - 1
                    prev_month_revenue = extract_revenue_by_year_month(revenue_data, prev_year, prev_month)
                    
                    if prev_month_revenue is not None and prev_month_revenue != 0:
                        mom = round((revenue / prev_month_revenue - 1) * 100, 2)
                        mom_row[month_label] = mom
                    else:
                        mom_row[month_label] = None
                    
                    # 計算 YoY（與去年同月比較）
                    last_year_revenue = extract_revenue_by_year_month(revenue_data, year - 1, month)
                    if last_year_revenue is not None and last_year_revenue != 0:
                        yoy = round((revenue / last_year_revenue - 1) * 100, 2)
                        yoy_row[month_label] = yoy
                    else:
                        yoy_row[month_label] = None
                else:
                    revenue_row[month_label] = None
                    mom_row[month_label] = None
                    yoy_row[month_label] = None
            
            rows.extend([revenue_row, mom_row, yoy_row])
        
        except Exception as e:
            logging.error(f"    錯誤: {stock_id} 營收總覽處理失敗 - {str(e)}")
            # 記錄詳細錯誤堆棧
            import traceback
            logging.error(f"    詳細錯誤:\n{traceback.format_exc()}")
            # 建立空白的三行
            revenue_row = {'代號': stock_id, '名稱': stock_name}
            mom_row = {'代號': '', '名稱': 'MoM(%)'}
            yoy_row = {'代號': '', '名稱': 'YoY(%)'}
            
            for year, month in months_list:
                month_label = f"{str(year)[-2:]}年{month}月"
                revenue_row[month_label] = None
                mom_row[month_label] = None
                yoy_row[month_label] = None
            
            rows.extend([revenue_row, mom_row, yoy_row])
    
    # 建立 DataFrame
    df = pd.DataFrame(rows, columns=columns)
    
    return df
