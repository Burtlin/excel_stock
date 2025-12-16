"""
股票數據處理主程式
整合營收、財務報表、EPS數據處理
"""
import logging
import os
import sys
from datetime import datetime

import pandas as pd
from FinMind.data import DataLoader

# 導入配置
from config import BASE_DIR

# 導入模組
from modules.logger import setup_logging, clean_old_logs
from modules.utils import process_info_data, format_percentage_columns
from modules.revenue import process_revenue_data, get_previous_three_months
from modules.financial import process_financial_data, process_eps_data




def process_stock(input_file='target.xlsx', output_file=None, revenue_sheet='月營收', financial_sheet='綜合損益表', eps_sheet='EPS'):
    """處理股票數據，營收、財務和 EPS 數據分別輸出到不同的 sheet"""
    # 初始化 logging
    setup_logging()
    clean_old_logs(days=7)
    
    logging.info("="*60)
    logging.info("開始處理股票數據")
    logging.info(f"輸入檔案: {input_file}")
    
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
    (last_month_year, last_month), (previous_month_year, previous_month), (previous_month_year2, previous_month2) = get_previous_three_months()
    
    # 計算去年同期
    yoy_year = last_month_year - 1
    
    total = len(df_base)
    for idx, row in df_base.iterrows():
        stock_id = row["代號"]
        logging.info(f"[{idx+1}/{total}] 處理中: {stock_id}")
        
        # 處理營收數據（寫入 df_revenue）
        process_revenue_data(api, df_revenue, idx, stock_id, last_month_year, last_month, previous_month_year, previous_month, previous_month_year2, previous_month2, yoy_year)
        
        # 處理綜合損益表數據（寫入 df_financial）
        try:
            process_financial_data(api, df_financial, idx, stock_id)
        except Exception as e:
            logging.error(f"  錯誤: {stock_id} 財務數據處理失敗 - {str(e)}")
        
        # 處理 EPS 數據（寫入 df_eps）
        try:
            process_eps_data(api, df_eps, idx, stock_id)
        except Exception as e:
            logging.error(f"  錯誤: {stock_id} EPS 數據處理失敗 - {str(e)}")
    
    logging.info("\n處理完成！")
    logging.info(f"\n營收數據:\n{df_revenue.head()}")
    logging.info(f"\n綜合損益表數據:\n{df_financial.head()}")
    logging.info(f"\nEPS數據:\n{df_eps.head()}")
    
    # 決定輸出檔案
    if output_file is None:
        output_file = input_file
    
    # 使用 openpyxl 保留原檔案的其他 sheet 和格式
    try:
        # 使用 ExcelWriter 將三個 DataFrame 分別寫入不同 sheet
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df_revenue.to_excel(writer, sheet_name=revenue_sheet, index=False)
            df_financial.to_excel(writer, sheet_name=financial_sheet, index=False)
            df_eps.to_excel(writer, sheet_name=eps_sheet, index=False)
            
            # 格式化百分比欄位
            format_percentage_columns(writer.sheets[revenue_sheet], df_revenue)
            format_percentage_columns(writer.sheets[financial_sheet], df_financial)
        
        logging.info(f"\n已更新並儲存至: {output_file}")
        logging.info(f"  - 營收數據: {revenue_sheet}")
        logging.info(f"  - 綜合損益表: {financial_sheet}")
        logging.info(f"  - EPS數據: {eps_sheet}")
    except Exception as e:
        logging.error(f"\n儲存檔案時發生錯誤: {str(e)}")
    
    logging.info("處理完成")
    logging.info("="*60 + "\n")
    
    return df_revenue, df_financial, df_eps


def main():
    """主程式進入點，增加錯誤處理"""
    try:
        args = sys.argv[1:]
        input_file = args[0] if len(args) > 0 else os.path.join(BASE_DIR, 'target.xlsx')
        output_file = args[1] if len(args) > 1 else input_file
        
        # 檢查檔案是否存在
        if not os.path.exists(input_file):
            print(f"錯誤: 找不到輸入檔案: {input_file}")
            print(f"請確保 {input_file} 檔案存在於程式目錄中")
            input("按 Enter 鍵離開...")
            return 1
        
        process_stock(input_file=input_file, output_file=output_file)
        return 0
    except Exception as e:
        print(f"程式執行失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        input("按 Enter 鍵離開...")
        return 1


if __name__ == '__main__':
    sys.exit(main())
