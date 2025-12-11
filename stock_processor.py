from datetime import datetime
import json
import os
import pandas as pd
from openpyxl import load_workbook
from FinMind.data import DataLoader
###########################################################################
# 名稱 相關功能
###########################################################################
def get_stock_name_mapping(api):
    """獲取股票代號與名稱的對應字典"""
    if os.path.exists('data/stock_info.json'):
        with open('data/stock_info.json', 'r', encoding='utf-8') as f:
            stock_dict = json.load(f)
        return stock_dict
    
    df = api.taiwan_stock_info()
    stock_dict = dict(zip(df['stock_id'], df['stock_name']))
    # save to json
    os.makedirs('data', exist_ok=True)
    with open('data/stock_info.json', 'w', encoding='utf-8') as f:
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

###########################################################################
# 營收 相關功能
###########################################################################

def get_stock_revenue_data(api, stock_id, start_date=None):
    """獲取股票營收數據"""
    if start_date is None:
        # 動態計算：取得兩年前的日期
        two_years_ago = datetime.now().replace(year=datetime.now().year - 2)
        start_date = two_years_ago.strftime('%Y-%m-%d')
    
    return api.taiwan_stock_month_revenue(
        stock_id=stock_id,
        start_date=start_date,
    )


def extract_revenue_by_year_month(revenue_data, target_year, target_month):
    """從營收數據中提取指定年月的營收"""
    for i in range(len(revenue_data) - 1, -1, -1):
        revenue_year = int(revenue_data.iloc[i]['revenue_year'])
        revenue_month = int(revenue_data.iloc[i]['revenue_month'])
        if revenue_year == target_year and revenue_month == target_month:
            return revenue_data.iloc[i]['revenue']
    return None

def get_previous_two_months():
    """取得上個月和上上個月的年份和月份"""
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
    
    return (last_month_year, last_month), (previous_month_year, previous_month)


def get_last_year_same_month(year, month):
    """取得去年同期的年份和月份"""
    return year - 1, month


def convert_to_million(value):
    """將數值從元轉換為百萬單位（取整）"""
    return round(value / 1000000) if value else None


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

##############################################################################
# EPS 相關功能
##############################################################################
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


def get_stock_financial_data(api, stock_id, start_date=None):
    """獲取股票財務報表數據"""
    if start_date is None:
        # 動態計算：取得兩年前的日期
        two_years_ago = datetime.now().replace(year=datetime.now().year - 2)
        start_date = two_years_ago.strftime('%Y-%m-%d')
    
    return api.taiwan_stock_financial_statement(
        stock_id=stock_id,
        start_date=start_date,
    )


def extract_value_by_date(financial_data, data_type, target_date):
    """從財務數據中提取指定類型和日期的數據（優化版）"""
    filtered_data = financial_data[
        (financial_data['type'] == data_type) & 
        (financial_data['date'] == target_date)
    ]
    return filtered_data.iloc[0]['value'] if not filtered_data.empty else None


def extract_eps_by_date(financial_data, target_date):
    """從財務數據中提取指定日期的 EPS"""
    return extract_value_by_date(financial_data, 'EPS', target_date)


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

############################################################################
# 主程序
############################################################################
def process_revenue_data(api, df, idx, stock_id, last_month_year, last_month, previous_month_year, previous_month, yoy_year):
    """處理單一股票的營收數據"""
    try:
        revenue_data = get_stock_revenue_data(api, stock_id)
        
        if revenue_data is None or revenue_data.empty:
            print(f"  警告: {stock_id} 無營收數據")
            return
    except Exception as e:
        print(f"  錯誤: {stock_id} 營收數據獲取失敗 - {str(e)}")
        return
    
    revenue_current = extract_revenue_by_year_month(revenue_data, last_month_year, last_month)
    revenue_previous = extract_revenue_by_year_month(revenue_data, previous_month_year, previous_month)
    revenue_yoy = extract_revenue_by_year_month(revenue_data, yoy_year, last_month)
    
    # 轉換為百萬單位
    revenue_current_million = convert_to_million(revenue_current)
    revenue_previous_million = convert_to_million(revenue_previous)
    
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
    df.at[idx, 'MoM(%)'] = mom
    df.at[idx, 'YoY(%)'] = yoy
    df.at[idx, f'{str(current_year)[-2:]}年累積營收(M)'] = ytd_revenue_million
    df.at[idx, '累積營收YoY(%)'] = ytd_yoy


def ensure_column_exists(df, column_name):
    """確保欄位存在，如果不存在則初始化"""
    if column_name not in df.columns:
        df[column_name] = None


def process_financial_data(api, df, idx, stock_id):
    """處理單一股票的綜合損益表（季營收、毛利率、累積營收）"""
    
    financial_data = get_stock_financial_data(api, stock_id)
    if financial_data is None or financial_data.empty:
        print(f"  警告: {stock_id} 無財務數據")
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
    
    financial_data = get_stock_financial_data(api, stock_id)
    if financial_data is None or financial_data.empty:
        print(f"  警告: {stock_id} 無財務數據")
        return
    
    # 處理 EPS（取得三季數據）
    try:
        eps_last, quarter_last, eps_prev, quarter_prev, eps_prev2, quarter_prev2 = get_last_three_season_data(financial_data, 'EPS')
    except Exception as e:
        print(f"  錯誤: {stock_id} EPS 提取失敗 - {str(e)}")
        eps_last = quarter_last = eps_prev = quarter_prev = eps_prev2 = quarter_prev2 = None
    
    # 計算今年累積 EPS
    try:
        ytd_eps = get_ytd_eps(financial_data)
    except Exception as e:
        print(f"  錯誤: {stock_id} YTD EPS 計算失敗 - {str(e)}")
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


def process_stock(input_file='target.xlsx', output_file=None, revenue_sheet='月營收', financial_sheet='綜合損益表', eps_sheet='EPS'):
    """處理股票數據，營收、財務和 EPS 數據分別輸出到不同的 sheet"""
    api = DataLoader()
    # api.login_by_token(api_token='token')
    # api.login(user_id='user_id', password='password')

    # 讀取第一個 sheet 取得股票代號
    df_base = pd.read_excel(input_file, sheet_name=0)
    df_base = df_base[['代號']].astype(int)
    
    # 創建三個 DataFrame：營收、綜合損益表、EPS
    df_revenue = df_base.copy()
    df_financial = df_base.copy()
    df_eps = df_base.copy()
    
    # 加入名稱欄位
    df_revenue = process_info_data(api, df_revenue)
    df_financial = process_info_data(api, df_financial)
    df_eps = process_info_data(api, df_eps)
    
    # 使用輔助函數計算時間
    (last_month_year, last_month), (previous_month_year, previous_month) = get_previous_two_months()
    
    # 計算去年同期
    yoy_year = last_month_year - 1
    
    # 初始化營收相關欄位
    current_year = datetime.now().year
    df_revenue[f'{last_month}月營收(M)'] = None
    df_revenue[f'{previous_month}月營收(M)'] = None
    df_revenue['MoM(%)'] = None
    df_revenue['YoY(%)'] = None
    df_revenue[f'{str(current_year)[-2:]}年累積營收(M)'] = None
    df_revenue['累積營收YoY(%)'] = None
    
    total = len(df_base)
    for idx, row in df_base.iterrows():
        stock_id = row["代號"]
        print(f"[{idx+1}/{total}] 處理中: {stock_id}")
        
        # 處理營收數據（寫入 df_revenue）
        process_revenue_data(api, df_revenue, idx, stock_id, last_month_year, last_month, previous_month_year, previous_month, yoy_year)
        
        # 處理綜合損益表數據（寫入 df_financial）
        try:
            process_financial_data(api, df_financial, idx, stock_id)
        except Exception as e:
            print(f"  錯誤: {stock_id} 財務數據處理失敗 - {str(e)}")
        
        # 處理 EPS 數據（寫入 df_eps）
        try:
            process_eps_data(api, df_eps, idx, stock_id)
        except Exception as e:
            print(f"  錯誤: {stock_id} EPS 數據處理失敗 - {str(e)}")
    
    print("\n處理完成！")
    print(f"\n營收數據:\n{df_revenue}")
    print(f"\n綜合損益表數據:\n{df_financial}")
    print(f"\nEPS數據:\n{df_eps}")
    
    # 決定輸出檔案
    if output_file is None:
        output_file = input_file
    
    # 使用 openpyxl 保留原檔案的其他 sheet 和格式
    try:
        # 載入原始工作簿
        book = load_workbook(input_file)
        
        # 使用 ExcelWriter 將三個 DataFrame 分別寫入不同 sheet
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df_revenue.to_excel(writer, sheet_name=revenue_sheet, index=False)
            df_financial.to_excel(writer, sheet_name=financial_sheet, index=False)
            df_eps.to_excel(writer, sheet_name=eps_sheet, index=False)
        
        print(f"\n已更新並儲存至: {output_file}")
        print(f"  - 營收數據: {revenue_sheet}")
        print(f"  - 綜合損益表: {financial_sheet}")
        print(f"  - EPS數據: {eps_sheet}")
    except Exception as e:
        print(f"\n儲存檔案時發生錯誤: {str(e)}")
    
    return df_revenue, df_financial, df_eps


if __name__ == '__main__':
    process_stock()

