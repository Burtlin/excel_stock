"""
單一股票分析 - 月營收年度比較
"""
import logging
import os
import sys
from datetime import datetime

import pandas as pd
from FinMind.data import DataLoader

from config import BASE_DIR
from modules.logger import setup_logging
from modules.utils import get_stock_name_mapping
from modules.revenue import get_stock_revenue_data
from modules.financial import get_stock_financial_data, get_last_season_month, get_previous_season_month, get_season_date, extract_value_by_date


def get_monthly_revenue_by_years(api, stock_id, years=3, use_cache=True):
    """取得指定股票近N年的月營收數據，整理成月份行格式，並新增MoM和YoY"""
    current_year = datetime.now().year
    
    # 計算起始年份
    start_year = current_year - years + 1
    start_date = f"{start_year}-01-01"
    
    # 取得營收數據（使用快取模組）
    if use_cache:
        logging.info(f"正在取得 {stock_id} 的營收數據（優先使用快取）...")
        revenue_data = get_stock_revenue_data(api, stock_id, start_date=start_date, use_cache=True)
    else:
        logging.info(f"正在從 API 取得 {stock_id} 的營收數據...")
        revenue_data = api.taiwan_stock_month_revenue(
            stock_id=stock_id,
            start_date=start_date
        )
        
        # 使用 API 抓取時自動更新快取
        if revenue_data is not None and not revenue_data.empty:
            from modules.cache import save_cache
            save_cache(stock_id, 'revenue', revenue_data)
            logging.info(f"✓ 已更新快取: {stock_id} 營收")
    
    if revenue_data is None or revenue_data.empty:
        logging.warning(f"查無 {stock_id} 的營收數據")
        return None
    
    # 輔助函數：取得指定年月的營收
    def get_revenue(year, month):
        data = revenue_data[
            (revenue_data['revenue_year'] == year) & 
            (revenue_data['revenue_month'] == month)
        ]
        if not data.empty:
            return data.iloc[0]['revenue']
        return None
    
    # 建立結果 DataFrame（橫縱對調：月份為行，年度為列）
    result_rows = []
    
    # 從 12 月到 1 月倒序排列
    for month in range(12, 0, -1):
        row = {'月份': f'{month}月'}
        
        # 每個年度作為一個欄位（由新到舊）
        for year in range(current_year, start_year - 1, -1):
            revenue = get_revenue(year, month)
            
            if revenue is not None:
                # 轉換為百萬單位
                revenue_million = round(revenue / 1000000)
                row[f'{year}年'] = revenue_million
            else:
                row[f'{year}年'] = None
        
        # 計算 MoM（Month-over-Month）- 與上個月比較
        current_revenue = get_revenue(current_year, month)
        prev_month = month - 1 if month > 1 else 12
        prev_year = current_year if month > 1 else current_year - 1
        prev_revenue = get_revenue(prev_year, prev_month)
        
        if current_revenue and prev_revenue and prev_revenue != 0:
            mom = round((current_revenue - prev_revenue) / prev_revenue * 100, 2)
            row['MoM(%)'] = mom
        else:
            row['MoM(%)'] = None
        
        # 計算 YoY（Year-over-Year）- 與去年同月比較
        last_year_revenue = get_revenue(current_year - 1, month)
        
        if current_revenue and last_year_revenue and last_year_revenue != 0:
            yoy = round((current_revenue - last_year_revenue) / last_year_revenue * 100, 2)
            row['YoY(%)'] = yoy
        else:
            row['YoY(%)'] = None
        
        result_rows.append(row)
    
    # 建立 DataFrame（年份降序：新的年份在先，最後加上 MoM 和 YoY）
    columns = ['月份'] + [f'{y}年' for y in range(current_year, start_year - 1, -1)] + ['MoM(%)', 'YoY(%)']
    df = pd.DataFrame(result_rows, columns=columns)
    
    return df


def get_financial_statement(api, stock_id, use_cache=True):
    """取得綜合損益表數據
    
    橫向：上一季、上上季、上上上季、上上上上季、今年累計、去年總共
    縱向：營業收入、毛利率、營益率、稅前淨利率、淨利率、EPS
    """
    # 取得財務數據
    if use_cache:
        logging.info(f"正在取得 {stock_id} 的財務數據（優先使用快取）...")
        financial_data = get_stock_financial_data(api, stock_id, use_cache=True)
    else:
        logging.info(f"正在從 API 取得 {stock_id} 的財務數據...")
        two_years_ago = datetime.now().replace(year=datetime.now().year - 2)
        start_date = two_years_ago.strftime('%Y-%m-%d')
        financial_data = api.taiwan_stock_financial_statement(
            stock_id=stock_id,
            start_date=start_date
        )
        
        # 使用 API 抓取時自動更新快取
        if financial_data is not None and not financial_data.empty:
            from modules.cache import save_cache
            save_cache(stock_id, 'financial', financial_data)
            logging.info(f"✓ 已更新快取: {stock_id} 財務")
    
    if financial_data is None or financial_data.empty:
        logging.warning(f"查無 {stock_id} 的財務數據")
        return None
    
    # 計算需要的季度
    current_year, q1_month = get_last_season_month()
    q2_year, q2_month = get_previous_season_month(current_year, q1_month)
    q3_year, q3_month = get_previous_season_month(q2_year, q2_month)
    q4_year, q4_month = get_previous_season_month(q3_year, q3_month)
    
    # 組合季度日期
    q1_date = get_season_date(current_year, q1_month)
    q2_date = get_season_date(q2_year, q2_month)
    q3_date = get_season_date(q3_year, q3_month)
    q4_date = get_season_date(q4_year, q4_month)
    
    # 季度名稱
    quarter_map = {3: 'Q1', 6: 'Q2', 9: 'Q3', 12: 'Q4'}
    q1_name = f"{str(current_year)[-2:]}{quarter_map[q1_month]}"
    q2_name = f"{str(q2_year)[-2:]}{quarter_map[q2_month]}"
    q3_name = f"{str(q3_year)[-2:]}{quarter_map[q3_month]}"
    q4_name = f"{str(q4_year)[-2:]}{quarter_map[q4_month]}"
    
    # 提取各項數據的函數
    def get_value(data_type, date):
        return extract_value_by_date(financial_data, data_type, date)
    
    # 計算今年累計（加總今年所有已公布季度）並取得季度數量
    def get_ytd(data_type):
        year = datetime.now().year
        data = financial_data[
            (financial_data['type'] == data_type) & 
            (financial_data['date'].str.startswith(str(year)))
        ]
        return data['value'].sum() if not data.empty else None
    
    def get_ytd_quarter_count():
        """計算今年已公布的季度數量"""
        year = datetime.now().year
        data = financial_data[
            (financial_data['type'] == 'Revenue') & 
            (financial_data['date'].str.startswith(str(year)))
        ]
        return len(data) if not data.empty else 0
    
    # 取得今年累計季度數
    ytd_quarters = get_ytd_quarter_count()
    ytd_column_name = f'今年累計({ytd_quarters}季)' if ytd_quarters > 0 else '今年累計'
    
    # 計算去年總共（加總去年四個季度）
    def get_last_year_total(data_type):
        year = datetime.now().year - 1
        data = financial_data[
            (financial_data['type'] == data_type) & 
            (financial_data['date'].str.startswith(str(year)))
        ]
        return data['value'].sum() if not data.empty else None
    
    # 建立資料行
    rows = []
    
    # 1. 營業收入（Revenue）- 單位：百萬
    revenue_q1 = get_value('Revenue', q1_date)
    revenue_q2 = get_value('Revenue', q2_date)
    revenue_q3 = get_value('Revenue', q3_date)
    revenue_q4 = get_value('Revenue', q4_date)
    revenue_ytd = get_ytd('Revenue')
    revenue_last_year = get_last_year_total('Revenue')
    
    rows.append({
        '項目': '營業收入',
        q1_name: round(revenue_q1 / 1000000) if revenue_q1 else None,
        q2_name: round(revenue_q2 / 1000000) if revenue_q2 else None,
        q3_name: round(revenue_q3 / 1000000) if revenue_q3 else None,
        q4_name: round(revenue_q4 / 1000000) if revenue_q4 else None,
        ytd_column_name: round(revenue_ytd / 1000000) if revenue_ytd else None,
        '去年總共': round(revenue_last_year / 1000000) if revenue_last_year else None
    })
    
    # 1-1. 營業收入 QoQ/YoY（Quarter-over-Quarter / Year-over-Year）- 單位：%
    def calc_qoq(current, previous):
        if current and previous and previous != 0:
            return round((current - previous) / previous * 100, 2)
        return None
    
    # 計算去年同期累計（用於YoY比較）
    def get_last_year_ytd(data_type):
        """計算去年同期累計（與今年相同季度數）"""
        last_year = datetime.now().year - 1
        
        # 取得去年數據
        data = financial_data[
            (financial_data['type'] == data_type) & 
            (financial_data['date'].str.startswith(str(last_year)))
        ]
        
        if data.empty:
            return None
        
        # 依日期排序並取前N季（N = 今年已公布季度數）
        data = data.sort_values('date').head(ytd_quarters)
        return data['value'].sum() if not data.empty else None
    
    revenue_last_year_ytd = get_last_year_ytd('Revenue')
    
    rows.append({
        '項目': 'QoQ/YoY',
        q1_name: calc_qoq(revenue_q1, revenue_q2),
        q2_name: calc_qoq(revenue_q2, revenue_q3),
        q3_name: calc_qoq(revenue_q3, revenue_q4),
        q4_name: None,  # 上上上上季沒有再上一季可比較
        ytd_column_name: calc_qoq(revenue_ytd, revenue_last_year_ytd),  # YoY: 今年累計 vs 去年同期累計
        '去年總共': None
    })
    
    # 2. 毛利率（GrossProfit / Revenue * 100）- 單位：%
    def calc_margin(gross_profit, revenue):
        if gross_profit and revenue and revenue != 0:
            return round(gross_profit / revenue * 100, 2)
        return None
    
    gross_q1 = get_value('GrossProfit', q1_date)
    gross_q2 = get_value('GrossProfit', q2_date)
    gross_q3 = get_value('GrossProfit', q3_date)
    gross_q4 = get_value('GrossProfit', q4_date)
    gross_ytd = get_ytd('GrossProfit')
    gross_last_year = get_last_year_total('GrossProfit')
    
    rows.append({
        '項目': '毛利率(%)',
        q1_name: calc_margin(gross_q1, revenue_q1),
        q2_name: calc_margin(gross_q2, revenue_q2),
        q3_name: calc_margin(gross_q3, revenue_q3),
        q4_name: calc_margin(gross_q4, revenue_q4),
        ytd_column_name: calc_margin(gross_ytd, revenue_ytd),
        '去年總共': calc_margin(gross_last_year, revenue_last_year)
    })
    
    # 3. 營益率（OperatingIncome / Revenue * 100）- 單位：%
    oi_q1 = get_value('OperatingIncome', q1_date)
    oi_q2 = get_value('OperatingIncome', q2_date)
    oi_q3 = get_value('OperatingIncome', q3_date)
    oi_q4 = get_value('OperatingIncome', q4_date)
    oi_ytd = get_ytd('OperatingIncome')
    oi_last_year = get_last_year_total('OperatingIncome')
    
    rows.append({
        '項目': '營益率(%)',
        q1_name: calc_margin(oi_q1, revenue_q1),
        q2_name: calc_margin(oi_q2, revenue_q2),
        q3_name: calc_margin(oi_q3, revenue_q3),
        q4_name: calc_margin(oi_q4, revenue_q4),
        ytd_column_name: calc_margin(oi_ytd, revenue_ytd),
        '去年總共': calc_margin(oi_last_year, revenue_last_year)
    })
    
    # 4. 稅前淨利率（PreTaxIncome / Revenue * 100）- 單位：%
    pti_q1 = get_value('PreTaxIncome', q1_date)
    pti_q2 = get_value('PreTaxIncome', q2_date)
    pti_q3 = get_value('PreTaxIncome', q3_date)
    pti_q4 = get_value('PreTaxIncome', q4_date)
    pti_ytd = get_ytd('PreTaxIncome')
    pti_last_year = get_last_year_total('PreTaxIncome')
    
    rows.append({
        '項目': '稅前淨利率(%)',
        q1_name: calc_margin(pti_q1, revenue_q1),
        q2_name: calc_margin(pti_q2, revenue_q2),
        q3_name: calc_margin(pti_q3, revenue_q3),
        q4_name: calc_margin(pti_q4, revenue_q4),
        ytd_column_name: calc_margin(pti_ytd, revenue_ytd),
        '去年總共': calc_margin(pti_last_year, revenue_last_year)
    })
    
    # 5. 淨利率（IncomeAfterTaxes / Revenue * 100）- 單位：%
    ni_q1 = get_value('IncomeAfterTaxes', q1_date)
    ni_q2 = get_value('IncomeAfterTaxes', q2_date)
    ni_q3 = get_value('IncomeAfterTaxes', q3_date)
    ni_q4 = get_value('IncomeAfterTaxes', q4_date)
    ni_ytd = get_ytd('IncomeAfterTaxes')
    ni_last_year = get_last_year_total('IncomeAfterTaxes')
    
    rows.append({
        '項目': '淨利率(%)',
        q1_name: calc_margin(ni_q1, revenue_q1),
        q2_name: calc_margin(ni_q2, revenue_q2),
        q3_name: calc_margin(ni_q3, revenue_q3),
        q4_name: calc_margin(ni_q4, revenue_q4),
        ytd_column_name: calc_margin(ni_ytd, revenue_ytd),
        '去年總共': calc_margin(ni_last_year, revenue_last_year)
    })
    
    # 6. EPS - 單位：元
    eps_q1 = get_value('EPS', q1_date)
    eps_q2 = get_value('EPS', q2_date)
    eps_q3 = get_value('EPS', q3_date)
    eps_q4 = get_value('EPS', q4_date)
    eps_ytd = get_ytd('EPS')
    eps_last_year = get_last_year_total('EPS')
    
    rows.append({
        '項目': 'EPS(元)',
        q1_name: round(eps_q1, 2) if eps_q1 else None,
        q2_name: round(eps_q2, 2) if eps_q2 else None,
        q3_name: round(eps_q3, 2) if eps_q3 else None,
        q4_name: round(eps_q4, 2) if eps_q4 else None,
        ytd_column_name: round(eps_ytd, 2) if eps_ytd else None,
        '去年總共': round(eps_last_year, 2) if eps_last_year else None
    })
    
    # 建立 DataFrame
    columns = ['項目', q1_name, q2_name, q3_name, q4_name, ytd_column_name, '去年總共']
    df = pd.DataFrame(rows, columns=columns)
    
    return df


def analyze_stock(stock_id, output_file=None, use_cache=True):
    """分析單一股票並輸出 Excel
    
    Args:
        stock_id: 股票代號
        output_file: 輸出檔案名稱
        use_cache: True=優先使用本地快取, False=強制從API抓取
    """
    # 初始化 logging
    setup_logging()
    
    logging.info("="*60)
    logging.info(f"開始分析股票: {stock_id}")
    logging.info(f"資料來源: {'本地快取/API' if use_cache else '強制API'}")
    
    api = DataLoader()
    
    # 取得股票名稱
    stock_dict = get_stock_name_mapping(api)
    stock_name = stock_dict.get(str(stock_id), "未知")
    logging.info(f"股票名稱: {stock_name}")
    
    # 取得月營收數據（近3年）
    df_revenue = get_monthly_revenue_by_years(api, stock_id, years=3, use_cache=use_cache)
    
    if df_revenue is None:
        logging.error("無法取得營收數據")
        return
    
    # 取得綜合損益表數據
    df_financial = get_financial_statement(api, stock_id, use_cache=use_cache)
    
    if df_financial is None:
        logging.warning("無法取得財務數據，僅輸出營收分析")
    
    # 決定輸出檔案名稱
    if output_file is None:
        output_file = f"{stock_id}_{stock_name}_分析.xlsx"
    
    # 輸出到 Excel
    try:
        import os
        
        # 判斷檔案是否存在，決定寫入模式
        file_exists = os.path.exists(output_file)
        mode = 'a' if file_exists else 'w'
        
        writer_kwargs = {'engine': 'openpyxl', 'mode': mode}
        if file_exists:
            writer_kwargs['if_sheet_exists'] = 'overlay'
        
        with pd.ExcelWriter(output_file, **writer_kwargs) as writer:
            df_revenue.to_excel(writer, sheet_name='月營收', index=False)
            
            # 格式化月營收的百分比欄位
            ws_revenue = writer.sheets['月營收']
            mom_col_idx = df_revenue.columns.get_loc('MoM(%)') + 1
            yoy_col_idx = df_revenue.columns.get_loc('YoY(%)') + 1
            
            for row_idx in range(2, ws_revenue.max_row + 1):
                # MoM 欄位
                cell_mom = ws_revenue.cell(row=row_idx, column=mom_col_idx)
                if cell_mom.value is not None and isinstance(cell_mom.value, (int, float)):
                    cell_mom.number_format = '0.00%'
                    cell_mom.value = cell_mom.value / 100
                
                # YoY 欄位
                cell_yoy = ws_revenue.cell(row=row_idx, column=yoy_col_idx)
                if cell_yoy.value is not None and isinstance(cell_yoy.value, (int, float)):
                    cell_yoy.number_format = '0.00%'
                    cell_yoy.value = cell_yoy.value / 100
            
            if df_financial is not None:
                df_financial.to_excel(writer, sheet_name='綜合損益表', index=False)
                
                # 格式化百分比欄位
                from modules.utils import format_percentage_columns
                ws = writer.sheets['綜合損益表']
                percentage_rows = ['毛利率(%)', '營益率(%)', '稅前淨利率(%)', '淨利率(%)', 'QoQ/YoY']
                
                for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1), start=2):
                    cell_value = row[0].value
                    if cell_value in percentage_rows:
                        for col_idx in range(2, ws.max_column + 1):
                            cell = ws.cell(row=row_idx, column=col_idx)
                            if cell.value is not None and isinstance(cell.value, (int, float)):
                                cell.number_format = '0.00%'
                                cell.value = cell.value / 100
        
        logging.info(f"\n已輸出至: {output_file}")
        logging.info(f"  - 月營收: {len(df_revenue)} 個月份數據")
        if df_financial is not None:
            logging.info(f"  - 綜合損益表: {len(df_financial)} 個財務指標")
    except Exception as e:
        logging.error(f"儲存檔案時發生錯誤: {str(e)}")
    
    logging.info("分析完成")
    logging.info("="*60 + "\n")


def main():
    """主程式進入點"""
    try:
        stock_id = input("請輸入股票編號: ")
        if stock_id.strip() == "":
            print("請輸入有效的股票編號")
            print("使用方式: python stock_analysis.py <股票代號> [選項]")
            print("\n選項:")
            print("  --no-cache    強制從API抓取，不使用本地快取")
            print("  -o <檔名>     指定輸出檔案名稱")
            print("\n範例:")
            print("  python stock_analysis.py")
            print("  python stock_analysis.py --no-cache")
            print("  python stock_analysis.py -o 台積電分析.xlsx")
            return 1
        
        use_cache = '--no-cache' not in sys.argv
        
        # 處理輸出檔名
        output_file = None
        if '-o' in sys.argv:
            o_index = sys.argv.index('-o')
            if o_index + 1 < len(sys.argv):
                output_file = sys.argv[o_index + 1]
        
        analyze_stock(stock_id, output_file, use_cache=use_cache)
        return 0
        
    except Exception as e:
        print(f"程式執行失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
