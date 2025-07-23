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
                    ws = wb.active
                    ws.title = '对账明细表'
                    
                    # 设置页面布局
                    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                    ws.page_setup.paperSize = ws.PAPERSIZE_A4
                    ws.page_setup.fitToPage = True
                    ws.page_setup.fitToHeight = 0
                    ws.page_setup.fitToWidth = 1
                    
                    # 设置页脚
                    ws.oddFooter.center.text = f'第 &P 页，共 &N 页'
                    
                    # 设置页边距（单位：英寸）
                    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5, header=0.3, footer=0.3)
                    
                    # 设置列宽
                    column_widths = {
                        'A': 15,  # 收货单号
                        'B': 12,  # 收货日期
                        'C': 30,  # 商品名称
                        'D': 10,  # 实收数量
                        'E': 10,  # 基本单位
                        'F': 10,  # 单价
                        'G': 12,  # 小计金额
                        'H': 10,  # 税额
                        'I': 8,   # 税率
                        'J': 12,  # 小计价税
                        'K': 15   # 部门
                    }
                    
                    for col, width in column_widths.items():
                        ws.column_dimensions[col].width = width
                    
                    # 设置标题
                    ws.merge_cells('A1:K1')
                    title_cell = ws['A1']
                    title_cell.value = f'{supplier_name} 对账明细表'
                    title_cell.font = Font(name='宋体', size=16, bold=True)
                    title_cell.alignment = Alignment(horizontal='center', vertical='center')
                    ws.row_dimensions[1].height = 30
                    
                    # 设置空白行
                    ws.merge_cells('A2:K2')
                    ws.row_dimensions[2].height = 10
                    
                    # 设置表头
                    headers = ['收货单号', '收货日期', '商品名称', '实收数量', '基本单位', '单价', '小计金额', '税额', '税率', '小计价税', '部门']
                    header_row = 3
                    for col_idx, header in enumerate(headers, 1):
                        cell = ws.cell(row=header_row, column=col_idx, value=header)
                        cell.font = Font(name='宋体', size=11, bold=True)
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        cell.fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
                        
                        # 设置边框
                        thin_border = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin')
                        )
                        cell.border = thin_border
                    
                    ws.row_dimensions[header_row].height = 20
                    
                    # 写入数据
                    start_row = header_row + 1
                    for idx, row in supplier_data_with_summary.iterrows():
                        row_num = start_row + idx
                        
                        # 设置行高
                        ws.row_dimensions[row_num].height = 20
                        
                        # 设置斑马线效果
                        fill = None
                        if idx % 2 == 0 and idx < len(supplier_data_with_summary) - 1:  # 偶数行添加浅灰色背景，但不包括合计行
                            fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
                        
                        # 如果是负数金额的行，使用浅红色背景
                        if row['小计价税'] < 0:
                            fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
                        
                        # 写入单元格数据
                        for col_idx, col_name in enumerate(headers, 1):
                            cell = ws.cell(row=row_num, column=col_idx, value=row[col_name])
                            
                            # 设置单元格格式
                            cell.font = Font(name='宋体', size=10)
                            
                            # 设置对齐方式
                            if col_name in ['收货单号', '收货日期', '基本单位', '税率', '部门']:
                                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            elif col_name in ['实收数量', '单价', '小计金额', '税额', '小计价税']:
                                cell.alignment = Alignment(horizontal='right', vertical='center')
                                
                                # 设置数字格式
                                if col_name == '单价':
                                    cell.number_format = '#,##0.00'
                                elif col_name == '税率':
                                    if pd.notna(row[col_name]) and row[col_name] != '':
                                        cell.number_format = '0%'
                                elif col_name in ['实收数量']:
                                    cell.number_format = '#,##0.000'
                                elif col_name in ['小计金额', '税额', '小计价税']:
                                    cell.number_format = '#,##0.00'
                            else:
                                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                            
                            # 应用填充
                            if fill:
                                cell.fill = fill
                            
                            # 设置边框
                            thin_border = Border(
                                left=Side(style='thin'),
                                right=Side(style='thin'),
                                top=Side(style='thin'),
                                bottom=Side(style='thin')
                            )
                            cell.border = thin_border
                            
                            # 合计行加粗
                            if idx == len(supplier_data_with_summary) - 1:  # 最后一行是合计行
                                cell.font = Font(name='宋体', size=10, bold=True)
                                cell.fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
                    
                    # 设置重复打印的行
                    ws.print_title_rows = f'1:{header_row}'
                    
                    # 创建商品汇总表
                    self.create_article_summary(wb, supplier_data, supplier_name)
                    
                    # 保存Excel文件
                    output_filename = os.path.join(year_month_dir, f'{supplier_name}_对账明细表_{year_month}.xlsx')
                    wb.save(output_filename)
                    logging.info(f'已生成供应商对账单: {output_filename}')
                    self.progress_signal.emit(f'已生成供应商对账单: {output_filename}')
            
            # 创建备份
            self.create_backup(final_df)
            
            self.finished_signal.emit(True, f'处理完成，共生成{current_supplier}个供应商对账单')
            
        except Exception as e:
            logging.error(f'处理过程中发生错误: {str(e)}', exc_info=True)
            self.progress_signal.emit(f'处理过程中发生错误: {str(e)}')
            self.finished_signal.emit(False, f'处理失败: {str(e)}')
    
    def create_article_summary(self, wb, supplier_data, supplier_name):
        """创建商品汇总表"""
        # 按商品名称分组统计
        article_summary = supplier_data.groupby('商品名称').agg({
            '实收数量': 'sum',
            '基本单位': 'first',
            '单价': 'mean',  # 使用平均单价
            '小计金额': 'sum',
            '小计价税': 'sum'
        }).reset_index()
        
        # 按总数量降序排序
        article_summary = article_summary.sort_values('实收数量', ascending=False)
        
        # 添加合计行
        summary_row = pd.DataFrame([{
            '商品名称': '合计',
            '实收数量': '',
            '基本单位': '',
            '单价': '',
            '小计金额': article_summary['小计金额'].sum(),
            '小计价税': article_summary['小计价税'].sum()
        }])
        
        article_summary = pd.concat([article_summary, summary_row], ignore_index=True)
        
        # 创建新的工作表
        ws = wb.create_sheet(title='Article_Summary')
        
        # 设置页面布局
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToHeight = 0
        ws.page_setup.fitToWidth = 1
        
        # 设置页脚
        ws.oddFooter.center.text = f'第 &P 页，共 &N 页'
        
        # 设置页边距（单位：英寸）
        ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5, header=0.3, footer=0.3)
        
        # 设置列宽
        column_widths = {
            'A': 40,  # 商品名称
            'B': 12,  # 实收数量
            'C': 10,  # 基本单位
            'D': 12,  # 单价
            'E': 15,  # 小计金额
            'F': 15   # 小计价税
        }
        
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width
        
        # 设置标题
        ws.merge_cells('A1:F1')
        title_cell = ws['A1']
        title_cell.value = f'{supplier_name} 商品汇总表'
        title_cell.font = Font(name='宋体', size=16, bold=True)
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 30
        
        # 设置空白行
        ws.merge_cells('A2:F2')
        ws.row_dimensions[2].height = 10
        
        # 设置表头
        headers = ['商品名称', '实收数量', '基本单位', '单价', '小计金额', '小计价税']
        header_row = 3
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=header)
            cell.font = Font(name='宋体', size=11, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
            
            # 设置边框
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            cell.border = thin_border
        
        ws.row_dimensions[header_row].height = 20
        
        # 写入数据
        start_row = header_row + 1
        for idx, row in article_summary.iterrows():
            row_num = start_row + idx
            
            # 设置行高
            ws.row_dimensions[row_num].height = 20
            
            # 设置斑马线效果
            fill = None
            if idx % 2 == 0 and idx < len(article_summary) - 1:  # 偶数行添加浅灰色背景，但不包括合计行
                fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
            
            # 写入单元格数据
            for col_idx, col_name in enumerate(headers, 1):
                cell = ws.cell(row=row_num, column=col_idx, value=row[col_name])
                
                # 设置单元格格式
                cell.font = Font(name='宋体', size=10)
                
                # 设置对齐方式
                if col_name in ['基本单位']:
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                elif col_name in ['实收数量', '单价', '小计金额', '小计价税']:
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    
                    # 设置数字格式
                    if col_name == '单价':
                        cell.number_format = '#,##0.00'
                    elif col_name == '实收数量':
                        cell.number_format = '#,##0.000'
                    elif col_name in ['小计金额', '小计价税']:
                        cell.number_format = '#,##0.00'
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                
                # 应用填充
                if fill:
                    cell.fill = fill
                
                # 设置边框
                thin_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                cell.border = thin_border
                
                # 合计行加粗
                if idx == len(article_summary) - 1:  # 最后一行是合计行
                    cell.font = Font(name='宋体', size=10, bold=True)
                    cell.fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
        
        # 设置重复打印的行
        ws.print_title_rows = f'1:{header_row}'
    
    def create_backup(self, final_df):
        """创建备份文件"""
        # 创建备份文件夹
        if not os.path.exists('bak'):
            os.makedirs('bak')
            logging.info('创建备份文件夹')
        
        # 保存备份数据
        backup_filename = os.path.join('bak', f'data_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
        final_df.to_excel(backup_filename, index=False)
        logging.info(f'已创建备份文件: {backup_filename}')
        self.progress_signal.emit(f'已创建备份文件: {backup_filename}')


class QTextEditLogger(logging.Handler):
    def __init__(self, parent):
        super().__init__()
        self.widget = QTextEdit(parent)
        self.widget.setReadOnly(True)
        self.widget.setLineWrapMode(QTextEdit.NoWrap)
        self.widget.setStyleSheet(
            "background-color: #f8f8f8; color: #333333; font-family: 'Courier New'; font-size: 10pt;"
        )
    
    def emit(self, record):
        msg = self.format(record)
        self.widget.append(msg)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.last_directory = None
        self.file_list = []
        self.initUI()
        self.centerOnScreen()
    
    def centerOnScreen(self):
        """将窗口居中显示"""
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) // 2, (screen.height() - size.height()) // 2)
    
    def initUI(self):
        """初始化UI"""
        # 设置窗口标题和图标
        self.setWindowTitle('MC 对账明细工具')
        self.setWindowIcon(QIcon(':/icon/app_icon.png'))
        
        # 设置窗口大小
        self.resize(800, 600)
        self.centerOnScreen()
        
        # 创建主窗口部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        
        # 文件选择区域
        file_group = QFrame()
        file_group.setFrameShape(QFrame.StyledPanel)
        file_group.setStyleSheet(
            "QFrame {background-color: #ffffff; border: 1px solid #dddddd; border-radius: 5px;}"
        )
        file_layout = QVBoxLayout(file_group)
        
        file_header = QLabel('文件选择')
        file_header.setStyleSheet(
            "font-weight: bold; font-size: 14px; color: #333333; padding: 5px;"
        )
        file_layout.addWidget(file_header)
        
        # 文件列表
        self.file_list_widget = QListWidget()
        self.file_list_widget.setStyleSheet(
            "background-color: #f8f8f8; border: 1px solid #dddddd; border-radius: 3px;"
        )
        file_layout.addWidget(self.file_list_widget)
        
        # 文件操作按钮
        file_buttons_layout = QHBoxLayout()
        
        self.add_files_button = QPushButton('添加文件')
        self.add_files_button.setStyleSheet(
            "QPushButton {background-color: #4CAF50; color: white; border-radius: 3px; padding: 5px 15px;}"
            "QPushButton:hover {background-color: #45a049;}"
            "QPushButton:pressed {background-color: #3d8b40;}"
        )
        self.add_files_button.clicked.connect(self.selectFiles)
        file_buttons_layout.addWidget(self.add_files_button)
        
        self.clear_files_button = QPushButton('清空列表')
        self.clear_files_button.setStyleSheet(
            "QPushButton {background-color: #f44336; color: white; border-radius: 3px; padding: 5px 15px;}"
            "QPushButton:hover {background-color: #d32f2f;}"
            "QPushButton:pressed {background-color: #b71c1c;}"
        )
        self.clear_files_button.clicked.connect(self.clearFiles)
        file_buttons_layout.addWidget(self.clear_files_button)
        
        file_layout.addLayout(file_buttons_layout)
        main_layout.addWidget(file_group, 2)
        
        # 进度显示区域
        progress_group = QFrame()
        progress_group.setFrameShape(QFrame.StyledPanel)
        progress_group.setStyleSheet(
            "QFrame {background-color: #ffffff; border: 1px solid #dddddd; border-radius: 5px;}"
        )
        progress_layout = QVBoxLayout(progress_group)
        
        progress_header = QLabel('处理进度')
        progress_header.setStyleSheet(
            "font-weight: bold; font-size: 14px; color: #333333; padding: 5px;"
        )
        progress_layout.addWidget(progress_header)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet(
            "QProgressBar {border: 1px solid #dddddd; border-radius: 3px; text-align: center;}"
            "QProgressBar::chunk {background-color: #2196F3;}"
        )
        self.progress_bar.setRange(0, 0)  # 设置为忙碌状态
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setVisible(False)
        progress_layout.addWidget(self.progress_bar)
        
        # 开始处理按钮
        self.start_button = QPushButton('开始处理')
        self.start_button.setStyleSheet(
            "QPushButton {background-color: #2196F3; color: white; border-radius: 3px; padding: 8px 20px; font-size: 12px;}"
            "QPushButton:hover {background-color: #0b7dda;}"
            "QPushButton:pressed {background-color: #0a69b7;}"
            "QPushButton:disabled {background-color: #cccccc; color: #666666;}"
        )
        self.start_button.clicked.connect(self.startProcess)
        progress_layout.addWidget(self.start_button, alignment=Qt.AlignCenter)
        
        main_layout.addWidget(progress_group, 1)
        
        # 日志显示区域
        log_group = QFrame()
        log_group.setFrameShape(QFrame.StyledPanel)
        log_group.setStyleSheet(
            "QFrame {background-color: #ffffff; border: 1px solid #dddddd; border-radius: 5px;}"
        )
        log_layout = QVBoxLayout(log_group)
        
        log_header = QLabel('处理日志')
        log_header.setStyleSheet(
            "font-weight: bold; font-size: 14px; color: #333333; padding: 5px;"
        )
        log_layout.addWidget(log_header)
        
        # 日志文本框
        self.log_handler = QTextEditLogger(self)
        self.log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)
        log_layout.addWidget(self.log_handler.widget)
        
        main_layout.addWidget(log_group, 3)
        
        # 版权信息
        copyright_label = QLabel('© 2023 MC Recon Tool. All rights reserved.')
        copyright_label.setStyleSheet(
            "color: #999999; font-size: 10px; padding: 5px;"
        )
        copyright_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(copyright_label)
        
        # 设置整体样式
        self.setStyleSheet(
            "QMainWindow {background-color: #f0f0f0;}"
            "QLabel {font-family: 'Microsoft YaHei', 'SimHei', sans-serif; color: #333333;}"
            "QTextEdit {font-family: 'Courier New'; font-size: 10pt;}"
            "QListWidget {font-family: 'Microsoft YaHei', 'SimHei', sans-serif; font-size: 10pt;}"
            "QListWidget::item {padding: 5px;}"
            "QListWidget::item:selected {background-color: #e0e0e0; color: #000000;}"
        )
    
    def selectFiles(self):
        """选择文件"""
        options = QFileDialog.Options()
        start_dir = self.last_directory if self.last_directory else os.path.expanduser("~")
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件", start_dir,
            "Excel Files (*.xlsx *.xls);;All Files (*)", options=options
        )
        
        if files:
            self.last_directory = os.path.dirname(files[0])
            
            # 添加文件到列表，避免重复
            for file in files:
                if file not in self.file_list:
                    self.file_list.append(file)
            
            self.updateFileList()
    
    def clearFiles(self):
        """清空文件列表"""
        self.file_list = []
        self.file_list_widget.clear()
    
    def updateFileList(self):
        """更新文件列表显示"""
        self.file_list_widget.clear()
        for file in self.file_list:
            item = QListWidgetItem(os.path.basename(file))
            item.setToolTip(file)
            self.file_list_widget.addItem(item)
    
    def startProcess(self):
        """开始处理数据"""
        if not self.file_list:
            QMessageBox.warning(self, "警告", "请先选择要处理的文件！")
            return
        
        # 禁用按钮，显示进度条
        self.start_button.setEnabled(False)
        self.add_files_button.setEnabled(False)
        self.clear_files_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        
        # 清空日志显示
        self.log_handler.widget.clear()
        
        # 创建并启动处理线程
        self.process_thread = DataProcessThread(self.file_list)
        self.process_thread.progress_signal.connect(self.updateProgress)
        self.process_thread.finished_signal.connect(self.processFinished)
        self.process_thread.start()
    
    def updateProgress(self, message):
        """更新进度信息"""
        # 滚动到底部
        self.log_handler.widget.moveCursor(self.log_handler.widget.textCursor().End)
    
    def processFinished(self, success, message):
        """处理完成后的操作"""
        # 恢复按钮状态，隐藏进度条
        self.start_button.setEnabled(True)
        self.add_files_button.setEnabled(True)
        self.clear_files_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if success:
            QMessageBox.information(self, "处理完成", message)
            
            # 询问是否打开输出文件夹
            reply = QMessageBox.question(self, '处理完成', 
                                       '是否打开输出文件夹？',
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            
            if reply == QMessageBox.Yes:
                output_dir = os.path.abspath('供应商对账明细')
                
                # 跨平台打开文件夹
                if sys.platform == 'win32':
                    os.startfile(output_dir)
                elif sys.platform == 'darwin':  # macOS
                    import subprocess
                    subprocess.Popen(['open', output_dir])
                else:  # linux
                    import subprocess
                    subprocess.Popen(['xdg-open', output_dir])
        else:
            QMessageBox.critical(self, "处理失败", message)
    
    def resetUI(self):
        """重置UI状态"""
        self.start_button.setEnabled(True)
        self.add_files_button.setEnabled(True)
        self.clear_files_button.setEnabled(True)
        self.progress_bar.setVisible(False)
    
    def showError(self, message):
        """显示错误信息"""
        QMessageBox.critical(self, "错误", message)
        self.resetUI()


def ensure_directories_exist():
    """确保必要的目录结构存在"""
    directories = ['logs', 'bak', '供应商对账明细']
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            logging.info(f'创建目录: {directory}')


def get_app_dir():
    """获取应用程序目录"""
    if getattr(sys, 'frozen', False):
        # 打包后的应用
        return os.path.dirname(sys.executable)
    else:
        # 开发环境
        return os.path.dirname(os.path.abspath(__file__))


def get_config_path():
    """获取配置文件路径"""
    app_dir = get_app_dir()
    return os.path.join(app_dir, 'config.ini')


def ensure_config_exists():
    """确保配置文件存在"""
    config_path = get_config_path()
    if not os.path.exists(config_path):
        config = configparser.ConfigParser()
        config['DEFAULT'] = {
            'LastDirectory': '',
            'BackupEnabled': 'True',
            'MaxBackupFiles': '10'
        }
        with open(config_path, 'w', encoding='utf-8') as f:
            config.write(f)
        logging.info(f'创建默认配置文件: {config_path}')
    return config_path


def check_expiration():
    """检查程序是否过期"""
    expiration_date = datetime(2025, 12, 31)  # 设置截止日期为2025年12月31日
    current_date = datetime.now()
    
    if current_date > expiration_date:
        return False, f"程序已于 {expiration_date.strftime('%Y年%m月%d日')} 过期，请联系管理员获取新版本。"
    
    days_left = (expiration_date - current_date).days
    if days_left <= 30:
        return True, f"程序将在 {days_left} 天后过期，请及时联系管理员获取新版本。"
    
    return True, ""


def main():
    # 配置日志
    if not os.path.exists('logs'):
        os.makedirs('logs')
    
    log_filename = os.path.join('logs', f'app_{datetime.now().strftime("%Y%m%d")}.log')
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    # 确保目录结构存在
    ensure_directories_exist()
    
    # 确保配置文件存在
    ensure_config_exists()
    
    # 创建Qt应用
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion风格，在所有平台上看起来一致
    
    # 设置应用图标
    app.setWindowIcon(QIcon(':/icon/app_icon.png'))
    
    # 在Windows平台上，设置任务栏图标
    if sys.platform == 'win32':
        import ctypes
        app_id = 'mc.recon.tool.1.0'  # 应用程序ID，可以是任意字符串
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    
    # 检查程序是否过期
    valid, message = check_expiration()
    if not valid:
        QMessageBox.critical(None, "程序已过期", message)
        return 1
    
    if message:  # 如果有警告消息（即将过期）
        QMessageBox.warning(None, "程序即将过期", message)
    
    # 创建并显示主窗口
    main_window = MainWindow()
    main_window.show()
    
    # 运行应用
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()