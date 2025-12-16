"""
日誌管理模組
"""
import logging
import os
from datetime import datetime, timedelta
from glob import glob

from config import LOGS_DIR, LOG_RETENTION_DAYS, LOG_FILENAME_FORMAT


def setup_logging():
    """設定logging，同時輸出到 console 和檔案"""
    # 建立logs目錄
    os.makedirs(LOGS_DIR, exist_ok=True)
    
    # 產生日誌檔案名稱（使用當前日期）
    log_filename = os.path.join(LOGS_DIR, datetime.now().strftime(LOG_FILENAME_FORMAT))
    
    # 設定logging格式
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler()  # 同時輸出到 console
        ]
    )
    
    return logging.getLogger(__name__)


def clean_old_logs(days=None):
    """刪除超過指定天數的舊日誌檔案"""
    if days is None:
        days = LOG_RETENTION_DAYS
    
    if not os.path.exists(LOGS_DIR):
        return
    
    cutoff_time = datetime.now() - timedelta(days=days)
    log_files = glob(os.path.join(LOGS_DIR, 'stock_processor_*.log'))
    
    deleted_count = 0
    for log_file in log_files:
        file_time = datetime.fromtimestamp(os.path.getmtime(log_file))
        if file_time < cutoff_time:
            try:
                os.remove(log_file)
                deleted_count += 1
            except Exception as e:
                logging.warning(f"無法刪除舊日誌: {log_file} - {str(e)}")
    
    if deleted_count > 0:
        logging.info(f"已刪除 {deleted_count} 個超過 {days} 天的舊日誌檔案")
