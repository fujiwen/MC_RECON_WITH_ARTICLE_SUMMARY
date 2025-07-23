import sys
import os
import pandas as pd
import numpy as np
import re
import logging
import configparser
from copy import copy
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side, Color
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins, PrintPageSetup
from openpyxl.worksheet.properties import PageSetupProperties
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QPushButton, QTextEdit, QProgressBar, QFrame,
                             QFileDialog, QMessageBox, QListWidget, QListWidgetItem)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QRect
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon
from PyQt5.QtWidgets import QDesktopWidget

class DataProcessThread(QThread):
    progress_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, input_files):
        super().__init__()
        self.input_files = input_files

    def format_mixed_text(self, text):
        if pd.isna(text):
            return text
        text = str(text)
        chinese_pattern = re.compile('[\u4e00-\u9fff]')
        match = chinese_pattern.search(text)
        if match:
            english_part = text[:match.start()].strip()
            chinese_part = text[match.start():].strip()
            if english_part and chinese_part:
                return f'{english_part}\n{chinese_part}'
        return text

    def extract_chinese(self, text):
        """提取文本中的中文字符"""
        if pd.isna(text):
            return text
        text = str(text)
        chinese_pattern = re.compile('[\u4e00-\u9fff]+')
        chinese_matches = chinese_pattern.findall(text)
        if chinese_matches:
            return ''.join(chinese_matches)
        return text

    def run(self):
        try:
            # 创建日志目录
            if not os.path.exists('logs'):
                os.makedirs('logs')
            
            # 配置日志
            log_filename = os.path.join('logs', f'process_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(log_filename, encoding='utf-8'),
                    logging.StreamHandler()
                ]
            )
            
            all_final_data = []
            
            for input_file in self.input_files:
                self.progress_signal.emit(f'开始读取文件：{os.path.basename(input_file)}')
                logging.info(f'开始读取文件：{input_file}')
                
                # 读取原始文件
                df = pd.read_excel(input_file, skiprows=8)
                logging.info(f'文件读取完成，共{len(df)}行数据')
                self.progress_signal.emit(f'文件读取完成，共{len(df)}行数据')
                
                # 获取收货单号的行索引
                receipt_rows = df[df['Unnamed: 0'].astype(str).str.match(r'^(RTS)?000\d+$', na=False)].index
                
                # 创建一个空的列表来存储所有明细数据
                all_details = []
                
                # 遍历每个收货单号之间的行
                total_receipts = len(receipt_rows)
                for i in range(total_receipts):
                    start_idx = receipt_rows[i]
                    end_idx = receipt_rows[i+1] if i < len(receipt_rows)-1 else len(df)
                    
                    receipt = df.loc[start_idx, 'Unnamed: 0']
                    supplier = df.loc[start_idx, 'Unnamed: 3']
                    date = df.loc[start_idx, 'Unnamed: 23']
                    
                    # 清理供应商名称和日期中的发票信息
                    if pd.notna(supplier):
                        supplier = re.sub(r'[（(].*[)）]|（专票.*|（普票.*|\s+专票.*|\s+普票.*|\d+%$', '', str(supplier)).strip()
                    
                    if pd.notna(date):
                        try:
                            date = pd.to_datetime(date)
                            if pd.notna(date):
                                date = date.strftime('%Y-%m-%d')
                        except:
                            date = None
                    
                    # 获取明细行（跳过收货单号行）
                    details = df.loc[start_idx+1:end_idx-1].copy()
                    
                    # 只保留非空行且不包含Page和Delivery Date的行
                    details = details[details['Unnamed: 0'].notna()]
                    details = details[~details['Unnamed: 0'].astype(str).str.contains('Page|Delivery Date', na=False)]
                    
                    if not details.empty:
                        details['收货单号'] = receipt
                        details['供应商名称'] = self.extract_chinese(supplier)
                        details['收货日期'] = date
                        details['商品名称'] = details['Unnamed: 0'].apply(self.format_mixed_text)
                        details['实收数量'] = details['Unnamed: 8']
                        details['基本单位'] = details['Unnamed: 9']
                        details['单价'] = details['Unnamed: 13']
                        details['小计金额'] = details['Unnamed: 25']
                        details['税额'] = details['Unnamed: 30']
                        details['税率'] = details['Unnamed: 30'] / details['Unnamed: 25']
                        details['小计价税'] = details['Unnamed: 34']
                        details['部门'] = details['Unnamed: 37'].apply(self.format_mixed_text)
                        
                        all_details.append(details[['收货单号', '收货日期', '商品名称', '实收数量', '基本单位',
                                                   '单价', '小计金额', '税额', '税率', '小计价税', '部门', '供应商名称']])
                    
                    progress = f'处理进度：{i+1}/{total_receipts}'
                    self.progress_signal.emit(progress)
                    logging.info(progress)
                
                # 合并所有明细数据
                if all_details:
                    file_df = pd.concat(all_details, ignore_index=True)
                    all_final_data.append(file_df)
                    logging.info(f'文件处理完成，共整理{len(file_df)}条记录')
                    self.progress_signal.emit(f'文件处理完成，共整理{len(file_df)}条记录')
            
            # 合并所有文件的数据
            final_df = pd.concat(all_final_data, ignore_index=True)
            logging.info(f'所有文件处理完成，共整理{len(final_df)}条记录')
            self.progress_signal.emit(f'所有文件处理完成，共整理{len(final_df)}条记录')
            
            # 创建供应商对账明细表文件夹
            if not os.path.exists('供应商对账明细'):
                os.makedirs('供应商对账明细')
                logging.info('创建供应商对账明细文件夹')
            
            # 按供应商名称分组并生成对账明细表
            total_suppliers = len(final_df['供应商名称'].unique())
            current_supplier = 0
            
            for supplier_name, supplier_data in final_df.groupby('供应商名称'):
                if pd.notna(supplier_name) and supplier_name.strip():
                    current_supplier += 1
                    self.progress_signal.emit(f'正在生成供应商对账单 ({current_supplier}/{total_suppliers}): {supplier_name}')
                    
                    # 按收货日期和收货单号排序
                    supplier_data = supplier_data.sort_values(['收货日期', '收货单号'])
                    
                    # 获取年月信息
                    first_date = pd.to_datetime(supplier_data['收货日期'].iloc[0])
                    year_month = first_date.strftime('%Y%m')
                    
                    # 创建年月目录
                    year_month_dir = os.path.join('供应商对账明细', year_month)
                    if not os.path.exists(year_month_dir):
                        os.makedirs(year_month_dir)
                    
                    # 计算合计金额
                    total_amount = supplier_data['小计价税'].sum()
                    
                    # 创建一个包含合计行的新数据框
                    summary_row = pd.DataFrame([{
                        '收货单号': '合计',
                        '收货日期': '',
                        '商品名称': '',
                        '实收数量': '',
                        '基本单位': '',
                        '单价': '',
                        '小计金额': supplier_data['小计金额'].sum(),
                        '税额': supplier_data['税额'].sum(),
                        '税率': '',
                        '小计价税': total_amount,
                        '部门': '',
                        '供应商名称': ''
                    }])
                    
                    supplier_data_with_summary = pd.concat([supplier_data, summary_row], ignore_index=True)
                    
                    # 创建新的Excel工作簿
                    wb = Workbook()