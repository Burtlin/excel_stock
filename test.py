from datetime import datetime

import pandas as pd
from FinMind.data import DataLoader


def get_stock_revenue_data(api, stock_id, start_date='2023-12-31'):
    """獲取股票營收數據"""
    return api.taiwan_stock_month_revenue(
        stock_id=stock_id,
        start_date=start_date,
    )


def extract_revenue_by_month(revenue_data, target_month):
    """從營收數據中提取指定月份的營收"""
    for i in range(len(revenue_data) - 1, -1, -1):
        revenue_month = int(revenue_data.iloc[i]['revenue_month'])
        if revenue_month == target_month:
            return revenue_data.iloc[i]['revenue']
    return None


def process_stock_revenue(input_file='target.xlsx', output_file=None):
    """處理股票營收數據"""
    api = DataLoader()
    # api.login_by_token(api_token='token')
    # api.login(user_id='user_id', password='password')

    df = pd.read_excel(input_file)
    last_month = datetime.now().month - 1
    
    for stock_id in df["代號"]:
        revenue_data = get_stock_revenue_data(api, stock_id)
        
        revenue_current = extract_revenue_by_month(revenue_data, last_month)
        revenue_previous = extract_revenue_by_month(revenue_data, last_month - 1)
        
        print(stock_id, last_month, revenue_current, revenue_previous)
        df.loc[df["代號"] == stock_id, f'{last_month}月營收'] = revenue_current
        df.loc[df["代號"] == stock_id, f'{last_month - 1}月營收'] = revenue_previous
    
    print(df)
    
    if output_file:
        df.to_excel(output_file, index=False)
    
    return df


if __name__ == '__main__':
    process_stock_revenue()
    # process_stock_revenue(output_file='output.xlsx')

