import sys
import os
import json
import re
from datetime import datetime
from pathlib import Path
from typing import List, Set, Dict

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QFileDialog, QListWidget, 
                            QLabel, QCheckBox, QMessageBox, QProgressBar,
                            QComboBox, QTextEdit, QSplitter, QFrame, QTabWidget,
                            QDialog, QDialogButtonBox)
from PyQt5.QtCore import Qt, pyqtSignal, QThread, QSize
from PyQt5.QtGui import QFont, QIcon, QPainter, QColor, QPen, QBrush, QPixmap, QImage, QPainter, QPageLayout
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageOps
import fitz  # PyMuPDF for PDF rendering
import tempfile
import subprocess
import platform
import shutil
import time

# 配置文件路径
CONFIG_LOG_FILE = "invoice_printer_config.json"
PRINT_LOG_FILE = "print_log.json"

class iOSStyleButton(QPushButton):
    """iOS风格按钮"""
    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QPushButton {
                background-color: #007AFF;
                color: white;
                border: none;
                border-radius: 12px;
                padding: 12px 20px;
                font-size: 16px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #0062CC;
            }
            QPushButton:pressed {
                background-color: #0052A3;
            }
            QPushButton:disabled {
                background-color: #C7C7CC;
                color: #8E8E93;
            }
        """)

class iOSStyleListWidget(QListWidget):
    """iOS风格列表控件"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QListWidget {
                background-color: white;
                border: 1px solid #E5E5EA;
                border-radius: 12px;
                padding: 10px;
                font-size: 14px;
            }
            QListWidget::item {
                padding: 10px;
                border-bottom: 1px solid #F0F0F0;
            }
            QListWidget::item:selected {
                background-color: #007AFF;
                color: white;
            }
        """)

class PrinterSelectionDialog(QDialog):
    """打印机选择对话框"""
    def __init__(self, available_printers: List[str], default_printer: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择打印机")
        self.available_printers = available_printers
        self.default_printer = default_printer
        self.selected_printer = None
        
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 标题
        title_label = QLabel("请选择要使用的打印机：")
        title_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        layout.addWidget(title_label)
        
        # 打印机列表
        self.printer_combo = QComboBox()
        self.printer_combo.addItems(self.available_printers)
        
        # 设置默认打印机为选中项
        if self.default_printer in self.available_printers:
            index = self.available_printers.index(self.default_printer)
            self.printer_combo.setCurrentIndex(index)
        
        self.printer_combo.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border: 1px solid #E5E5EA;
                border-radius: 8px;
                font-size: 14px;
            }
        """)
        layout.addWidget(self.printer_combo)
        
        # 按钮
        btn_layout = QHBoxLayout()
        
        ok_btn = QPushButton("确定")
        ok_btn.setStyleSheet("""
            QPushButton {
                background-color: #007AFF;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
                font-size: 14px;
                font-weight: 600;
            }
        """)
        ok_btn.clicked.connect(self.accept_selection)
        
        cancel_btn = QPushButton("取消")
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #C7C7CC;
                color: #1C1C1E;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
                font-size: 14px;
                font-weight: 600;
            }
        """)
        cancel_btn.clicked.connect(self.reject)
        
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
    
    def accept_selection(self):
        self.selected_printer = self.printer_combo.currentText()
        self.accept()

class OFDConverter:
    """OFD格式转换器"""
    @staticmethod
    def convert_to_pdf(ofd_path: str) -> str:
        """
        将OFD文件转换为PDF
        """
        try:
            # 方法1：尝试使用 ofdpy 库（如果已安装）
            try:
                from ofdpy import OFD
                ofd_doc = OFD(ofd_path)
                pdf_path = ofd_path.replace('.ofd', '.pdf')
                ofd_doc.to_pdf(pdf_path)
                return pdf_path
            except ImportError:
                pass
            
            # 方法2：创建一个包含原始文件信息的PDF作为占位符
            return OFDConverter.create_dummy_pdf(ofd_path)
            
        except Exception as e:
            print(f"OFD转换失败: {e}")
            return OFDConverter.create_dummy_pdf(ofd_path)
    
    @staticmethod
    def create_dummy_pdf(ofd_path: str) -> str:
        """创建一个包含原始文件信息的PDF"""
        import tempfile
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        
        temp_pdf = ofd_path.replace('.ofd', '_converted.pdf')
        
        c = canvas.Canvas(temp_pdf, pagesize=letter)
        width, height = letter
        
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, height - 50, "OFD文件转换")
        
        c.setFont("Helvetica", 12)
        c.drawString(50, height - 80, f"原始文件: {os.path.basename(ofd_path)}")
        c.drawString(50, height - 100, "注意：此文件是转换后的占位符")
        c.drawString(50, height - 120, "实际应用中应使用专业OFD转换工具")
        
        c.showPage()
        c.save()
        
        return temp_pdf

class DuplicateInvoiceDialog(QDialog):
    """重复发票确认对话框"""
    def __init__(self, duplicate_files: List[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("重复发票提醒")
        self.duplicate_files = duplicate_files
        self.result = None
        
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 标题
        title_label = QLabel("发现以下发票已打印过：")
        title_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        layout.addWidget(title_label)
        
        # 列表
        list_widget = QListWidget()
        for file in self.duplicate_files:
            list_widget.addItem(os.path.basename(file))
        layout.addWidget(list_widget)
        
        # 说明
        info_label = QLabel("请选择如何处理这些重复发票：")
        layout.addWidget(info_label)
        
        # 按钮组
        btn_layout = QHBoxLayout()
        
        all_yes_btn = QPushButton("全部是")
        all_yes_btn.setStyleSheet("""
            QPushButton {
                background-color: #007AFF;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
                font-size: 14px;
                font-weight: 600;
            }
        """)
        all_yes_btn.clicked.connect(lambda: self.accept_all(True))
        
        all_no_btn = QPushButton("全否")
        all_no_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF3B30;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
                font-size: 14px;
                font-weight: 600;
            }
        """)
        all_no_btn.clicked.connect(lambda: self.accept_all(False))
        
        cancel_btn = QPushButton("取消")
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #C7C7CC;
                color: #1C1C1E;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
                font-size: 14px;
                font-weight: 600;
            }
        """)
        cancel_btn.clicked.connect(self.reject)
        
        btn_layout.addWidget(all_yes_btn)
        btn_layout.addWidget(all_no_btn)
        btn_layout.addWidget(cancel_btn)
        
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
    
    def accept_all(self, accept: bool):
        self.result = accept
        self.accept()

class PrintWorker(QThread):
    """打印工作线程"""
    progress_updated = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str, int)  # 新增参数：打印份数
    
    def __init__(self, invoice_files: List[str], output_path: str, is_pdf: bool, selected_printer: str = None, 
                 page_layout: str = "1x2", copies_per_invoice: int = 2, parent=None):
        super().__init__(parent)
        self.invoice_files = invoice_files
        self.output_path = output_path
        self.is_pdf = is_pdf
        self.selected_printer = selected_printer  # 选择的打印机
        self.page_layout = page_layout  # 页面布局
        self.copies_per_invoice = copies_per_invoice  # 每张发票的份数
        
    def run(self):
        try:
            if self.is_pdf:
                success, printed_count = self.create_pdf()
            else:
                success, printed_count = self.print_directly_system()
            
            if success:
                self.finished.emit(True, "打印完成！", printed_count)
            else:
                self.finished.emit(False, "打印失败！", 0)
        except Exception as e:
            self.finished.emit(False, f"错误: {str(e)}", 0)
    
    def get_items_per_page(self):
        """根据布局获取每页项目数"""
        layout_map = {
            '1x1': 1,
            '1x2': 2,
            '1x3': 3,
            '2x2': 4,
            '2x3': 6,
            '2x4': 8
        }
        return layout_map.get(self.page_layout, 2)
    
    def create_pdf(self) -> tuple:
        """创建PDF文件"""
        try:
            c = canvas.Canvas(self.output_path, pagesize=A4)
            width, height = A4
            
            # 创建所有需要打印的发票列表（考虑份数）
            all_invoices_to_print = []
            for invoice_file in self.invoice_files:
                for copy_num in range(self.copies_per_invoice):
                    all_invoices_to_print.append(invoice_file)
            
            total_items = len(all_invoices_to_print)
            items_per_page = self.get_items_per_page()
            pages_needed = (total_items + items_per_page - 1) // items_per_page
            
            for page in range(pages_needed):
                self.progress_updated.emit(int((page + 1) / pages_needed * 100), 
                                         f"处理第 {page + 1}/{pages_needed} 页")
                
                # 计算每个发票的位置
                if self.page_layout == "1x1":
                    # 1x1: 整个页面一张发票
                    if page < total_items:
                        self.draw_invoice(c, all_invoices_to_print[page], 0, 0, width, height)
                
                elif self.page_layout == "1x2":
                    # 1x2: 纵向上下各一张
                    cell_height = height / 2
                    for i in range(2):
                        item_index = page * 2 + i
                        if item_index < total_items:
                            y = height - (i + 1) * cell_height
                            self.draw_invoice(c, all_invoices_to_print[item_index], 0, y, width, cell_height)
                
                elif self.page_layout == "1x3":
                    # 1x3: 纵向上中下各一张
                    cell_height = height / 3
                    for i in range(3):
                        item_index = page * 3 + i
                        if item_index < total_items:
                            y = height - (i + 1) * cell_height
                            self.draw_invoice(c, all_invoices_to_print[item_index], 0, y, width, cell_height)
                
                elif self.page_layout == "2x2":
                    # 2x2: 上下各2张（2行2列）
                    cell_width = width / 2
                    cell_height = height / 2
                    for row in range(2):
                        for col in range(2):
                            item_index = page * 4 + row * 2 + col
                            if item_index < total_items:
                                x = col * cell_width
                                y = height - (row + 1) * cell_height
                                self.draw_invoice(c, all_invoices_to_print[item_index], x, y, cell_width, cell_height)
                
                elif self.page_layout == "2x3":
                    # 2x3: 2行3列（每页6张）
                    cell_width = width / 3
                    cell_height = height / 2
                    for row in range(2):
                        for col in range(3):
                            item_index = page * 6 + row * 3 + col
                            if item_index < total_items:
                                x = col * cell_width
                                y = height - (row + 1) * cell_height
                                self.draw_invoice(c, all_invoices_to_print[item_index], x, y, cell_width, cell_height)
                
                elif self.page_layout == "2x4":
                    # 2x4: 2行4列（每页8张）
                    cell_width = width / 4
                    cell_height = height / 2
                    for row in range(2):
                        for col in range(4):
                            item_index = page * 8 + row * 4 + col
                            if item_index < total_items:
                                x = col * cell_width
                                y = height - (row + 1) * cell_height
                                self.draw_invoice(c, all_invoices_to_print[item_index], x, y, cell_width, cell_height)
                
                c.showPage()
            
            c.save()
            return True, len(self.invoice_files)
        except Exception as e:
            print(f"PDF创建错误: {e}")
            return False, 0
    
    def draw_invoice(self, c, invoice_path: str, x: float, y: float, 
                    width: float, height: float):
        """在指定位置绘制发票（不显示文件名）"""
        try:
            # 根据文件扩展名选择处理方式
            ext = os.path.splitext(invoice_path)[1].lower()
            
            if ext == '.ofd':
                # 处理OFD文件
                self.draw_ofd_invoice(c, invoice_path, x, y, width, height)
            elif ext in ['.pdf', '.png', '.jpg', '.jpeg', '.bmp']:
                # 处理PDF或图片文件
                self.draw_pdf_or_image_invoice(c, invoice_path, x, y, width, height)
            else:
                # 其他格式
                self.draw_text_only_invoice(c, invoice_path, x, y, width, height)
                
        except Exception as e:
            print(f"绘制发票错误 {invoice_path}: {e}")
            c.setFont("Helvetica", 10)
            c.setFillColorRGB(1, 0, 0)
            c.drawString(x + 20, y + height/2, f"错误: {os.path.basename(invoice_path)}")
            c.setFillColorRGB(0, 0, 0)
    
    def draw_ofd_invoice(self, c, ofd_path: str, x: float, y: float, 
                        width: float, height: float):
        """绘制OFD发票"""
        try:
            # 尝试转换为PDF
            converter = OFDConverter()
            pdf_path = converter.convert_to_pdf(ofd_path)
            
            # 使用PyMuPDF渲染PDF（超高分辨率）
            doc = fitz.open(pdf_path)
            page = doc[0]  # 获取第一页
            
            # 创建临时图像（超高分辨率 - 4x缩放）
            zoom = 4.0  # 超高缩放因子以获得极致清晰度
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)  # 禁用透明度以提高质量
            
            # 保存为临时PNG文件
            temp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            temp_img_path = temp_img.name
            temp_img.close()
            
            pix.save(temp_img_path)
            
            # 绘制图像（不显示标题）
            self.draw_image_to_canvas(c, temp_img_path, x, y, width, height)
            
            # 清理临时文件
            os.unlink(temp_img_path)
            if pdf_path != ofd_path and pdf_path.endswith('_converted.pdf'):
                os.unlink(pdf_path)
                
        except Exception as e:
            print(f"OFD渲染错误: {e}")
            self.draw_text_only_invoice(c, ofd_path, x, y, width, height)
    
    def draw_pdf_or_image_invoice(self, c, file_path: str, x: float, y: float, 
                                 width: float, height: float):
        """绘制PDF或图片发票"""
        try:
            ext = os.path.splitext(file_path)[1].lower()
            
            if ext == '.pdf':
                # 使用PyMuPDF渲染PDF（超高分辨率）
                doc = fitz.open(file_path)
                page = doc[0]  # 获取第一页
                
                # 创建临时图像（超高分辨率 - 4x缩放）
                zoom = 4.0  # 超高缩放因子
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)  # 禁用透明度
                
                # 保存为临时PNG文件
                temp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                temp_img_path = temp_img.name
                temp_img.close()
                
                pix.save(temp_img_path)
                
                # 绘制图像（不显示标题）
                self.draw_image_to_canvas(c, temp_img_path, x, y, width, height)
                
                # 清理临时文件
                os.unlink(temp_img_path)
                
            else:
                # 图片文件直接绘制
                self.draw_image_to_canvas(c, file_path, x, y, width, height)
                
        except Exception as e:
            print(f"PDF/图片渲染错误: {e}")
            self.draw_text_only_invoice(c, file_path, x, y, width, height)
    
    def draw_image_to_canvas(self, c, img_path: str, x: float, y: float, 
                           width: float, height: float):
        """将图像绘制到Canvas上（不显示文件名）"""
        try:
            # 加载图像
            img = Image.open(img_path)
            img_width, img_height = img.size
            
            # 计算缩放比例以适应区域
            scale_x = width / img_width
            scale_y = height / img_height  # 不留标题空间
            scale = min(scale_x, scale_y, 1.0)
            
            new_width = img_width * scale
            new_height = img_height * scale
            
            # 居中显示
            img_x = x + (width - new_width) / 2
            img_y = y + (height - new_height) / 2
            
            # 直接绘制图片，不添加任何文字
            c.drawImage(img_path, img_x, img_y, new_width, new_height)
            
        except Exception as e:
            print(f"图像绘制错误: {e}")
            c.setFont("Helvetica", 10)
            c.drawString(x + 20, y + height/2, "无法显示发票内容")
    
    def draw_text_only_invoice(self, c, file_path: str, x: float, y: float, 
                              width: float, height: float):
        """仅显示文本的发票"""
        try:
            # 显示基本信息
            filename = os.path.basename(file_path)
            c.setFont("Helvetica-Bold", 12)
            c.drawString(x + 20, y + height/2, f"发票文件: {filename}")
            c.setFont("Helvetica", 10)
            c.drawString(x + 20, y + height/2 - 20, "（内容无法渲染）")
            
        except Exception as e:
            print(f"文本绘制错误: {e}")
            c.setFont("Helvetica", 10)
            c.drawString(x + 20, y + height/2, f"错误: {filename}")
    
    def print_directly_system(self) -> tuple:
        """使用系统原生打印功能（终极修复版本）"""
        try:
            # 创建临时PDF文件用于系统打印
            temp_dir = os.path.expanduser("~/Documents") if os.path.exists(os.path.expanduser("~/Documents")) else os.getcwd()
            temp_pdf_path = os.path.join(temp_dir, f"invoice_print_{int(time.time())}.pdf")
            
            # 创建PDF文件（根据用户设置的布局）
            c = canvas.Canvas(temp_pdf_path, pagesize=A4)
            width, height = A4
            
            # 创建所有需要打印的发票列表（考虑份数）
            all_invoices_to_print = []
            for invoice_file in self.invoice_files:
                for copy_num in range(self.copies_per_invoice):
                    all_invoices_to_print.append(invoice_file)
            
            total_items = len(all_invoices_to_print)
            items_per_page = self.get_items_per_page()
            pages_needed = (total_items + items_per_page - 1) // items_per_page
            
            for page in range(pages_needed):
                # 计算每个发票的位置
                if self.page_layout == "1x1":
                    if page < total_items:
                        self.draw_invoice(c, all_invoices_to_print[page], 0, 0, width, height)
                
                elif self.page_layout == "1x2":
                    cell_height = height / 2
                    for i in range(2):
                        item_index = page * 2 + i
                        if item_index < total_items:
                            y = height - (i + 1) * cell_height
                            self.draw_invoice(c, all_invoices_to_print[item_index], 0, y, width, cell_height)
                
                elif self.page_layout == "1x3":
                    cell_height = height / 3
                    for i in range(3):
                        item_index = page * 3 + i
                        if item_index < total_items:
                            y = height - (i + 1) * cell_height
                            self.draw_invoice(c, all_invoices_to_print[item_index], 0, y, width, cell_height)
                
                elif self.page_layout == "2x2":
                    cell_width = width / 2
                    cell_height = height / 2
                    for row in range(2):
                        for col in range(2):
                            item_index = page * 4 + row * 2 + col
                            if item_index < total_items:
                                x = col * cell_width
                                y = height - (row + 1) * cell_height
                                self.draw_invoice(c, all_invoices_to_print[item_index], x, y, cell_width, cell_height)
                
                elif self.page_layout == "2x3":
                    cell_width = width / 3
                    cell_height = height / 2
                    for row in range(2):
                        for col in range(3):
                            item_index = page * 6 + row * 3 + col
                            if item_index < total_items:
                                x = col * cell_width
                                y = height - (row + 1) * cell_height
                                self.draw_invoice(c, all_invoices_to_print[item_index], x, y, cell_width, cell_height)
                
                elif self.page_layout == "2x4":
                    cell_width = width / 4
                    cell_height = height / 2
                    for row in range(2):
                        for col in range(4):
                            item_index = page * 8 + row * 4 + col
                            if item_index < total_items:
                                x = col * cell_width
                                y = height - (row + 1) * cell_height
                                self.draw_invoice(c, all_invoices_to_print[item_index], x, y, cell_width, cell_height)
                
                c.showPage()
            
            c.save()
            
            # 使用系统原生打印功能
            success = False
            
            if platform.system() == "Windows":
                # Windows系统 - 使用最可靠的方法：调用系统打印对话框
                try:
                    # 使用 win32api 调用系统打印对话框
                    import win32api
                    import win32print
                    
                    # 如果指定了打印机，使用特定打印机
                    if self.selected_printer:
                        printer_handle = win32print.OpenPrinter(self.selected_printer)
                        win32print.StartDocPrinter(printer_handle, 1, ("Invoice_Print", None, "RAW"))
                        win32print.StartPagePrinter(printer_handle)
                        
                        # 读取PDF文件并发送到打印机
                        with open(temp_pdf_path, 'rb') as f:
                            data = f.read()
                            win32print.WritePrinter(printer_handle, data)
                        
                        win32print.EndPagePrinter(printer_handle)
                        win32print.EndDocPrinter(printer_handle)
                        win32print.ClosePrinter(printer_handle)
                        success = True
                    else:
                        # 使用默认打印机
                        win32api.ShellExecute(0, "print", temp_pdf_path, None, ".", 0)
                        success = True
                except Exception as e:
                    print(f"Windows打印失败: {e}")
                    # 保留临时文件供用户手动打印
                    print(f"请手动打印临时文件: {temp_pdf_path}")
                    success = False
            
            elif platform.system() == "Darwin":  # macOS
                try:
                    # macOS使用lp命令打印
                    cmd = ['lp']
                    if self.selected_printer:
                        cmd.extend(['-d', self.selected_printer])
                    cmd.append(temp_pdf_path)
                    
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                    if result.returncode == 0:
                        success = True
                    else:
                        # 如果lp命令失败，尝试使用open命令
                        subprocess.run(['open', '-a', 'Preview', temp_pdf_path])
                        success = True
                except Exception as e:
                    print(f"macOS打印失败: {e}")
                    print(f"请手动打印临时文件: {temp_pdf_path}")
                    success = False
            
            else:  # Linux
                try:
                    # Linux使用lp命令打印
                    cmd = ['lp']
                    if self.selected_printer:
                        cmd.extend(['-d', self.selected_printer])
                    cmd.append(temp_pdf_path)
                    
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                    if result.returncode == 0:
                        success = True
                    else:
                        # 如果lp命令失败，尝试使用lpr命令
                        lpr_cmd = ['lpr']
                        if self.selected_printer:
                            lpr_cmd.extend(['-P', self.selected_printer])
                        lpr_cmd.append(temp_pdf_path)
                        subprocess.run(lpr_cmd)
                        success = True
                except Exception as e:
                    print(f"Linux打印失败: {e}")
                    print(f"请手动打印临时文件: {temp_pdf_path}")
                    success = False
            
            # 不立即删除临时文件，让用户有机会手动打印
            if success:
                # 只有在确认打印成功后才删除文件（等待10秒）
                time.sleep(10)
                try:
                    os.unlink(temp_pdf_path)
                except:
                    pass  # 如果文件还在使用中，忽略错误
                return True, len(self.invoice_files)
            else:
                # 打印失败时保留文件
                print(f"打印失败，临时文件已保存: {temp_pdf_path}")
                return False, 0
            
        except Exception as e:
            print(f"系统打印错误: {e}")
            return False, 0

class InvoicePrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ZZ发票打印助手")
        self.setGeometry(100, 100, 1000, 700)
        
        # 加载配置
        self.config = self.load_config()
        self.printed_invoices = self.load_print_log()
        
        # 当前发票文件列表
        self.invoice_files = []
        self.file_info = {}  # 存储文件信息
        
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        # 主窗口样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F2F2F7;
            }
        """)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # 标题
        title_label = QLabel("ZZ发票打印助手")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #1C1C1E;
                padding: 10px;
            }
        """)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 分割器：左侧操作面板，右侧文件列表
        splitter = QSplitter(Qt.Horizontal)
        
        # 左侧：操作面板（占1份）
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        # 操作按钮
        btn_layout = QVBoxLayout()
        btn_layout.setSpacing(10)
        
        self.add_files_btn = iOSStyleButton("添加发票文件")
        self.add_folder_btn = iOSStyleButton("从文件夹导入")
        self.remove_selected_btn = iOSStyleButton("移除选中文件")
        self.clear_all_btn = iOSStyleButton("清空列表")
        
        # 退出按钮改为浅红色
        self.exit_btn = QPushButton("退出程序")
        self.exit_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF6B6B;  /* 浅红色 */
                color: white;
                border: none;
                border-radius: 12px;
                padding: 12px 20px;
                font-size: 16px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #FF5252;  /* 稍深一点的红色 */
            }
            QPushButton:pressed {
                background-color: #FF3B30;  /* 更深的红色 */
            }
            QPushButton:disabled {
                background-color: #C7C7CC;
                color: #8E8E93;
            }
        """)
        
        for btn in [self.add_files_btn, self.add_folder_btn, 
                   self.remove_selected_btn, self.clear_all_btn]:
            btn_layout.addWidget(btn)
        btn_layout.addWidget(self.exit_btn)
        
        left_layout.addLayout(btn_layout)
        
        # 打印选项
        options_frame = QFrame()
        options_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 12px;
                padding: 15px;
            }
        """)
        options_layout = QVBoxLayout(options_frame)
        
        options_layout.addWidget(QLabel("打印选项:"))
        
        self.print_mode_combo = QComboBox()
        self.print_mode_combo.addItems(["打印为PDF", "直接打印"])
        self.print_mode_combo.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border: 1px solid #E5E5EA;
                border-radius: 8px;
                font-size: 14px;
            }
        """)
        options_layout.addWidget(self.print_mode_combo)
        
        # 页面布局选项（明确说明布局方式）
        options_layout.addWidget(QLabel("页面布局:"))
        self.page_layout_combo = QComboBox()
        self.page_layout_combo.addItems([
            "1×1 (每页1张)",
            "1×2 (每页2张，纵向上下)", 
            "1×3 (每页3张，纵向上中下)",
            "2×2 (每页4张，2行2列)",
            "2×3 (每页6张，2行3列)",
            "2×4 (每页8张，2行4列)"
        ])
        self.page_layout_combo.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border: 1px solid #E5E5EA;
                border-radius: 8px;
                font-size: 14px;
            }
        """)
        options_layout.addWidget(self.page_layout_combo)
        
        # 每张发票打印份数
        options_layout.addWidget(QLabel("每张发票打印份数:"))
        self.copies_combo = QComboBox()
        self.copies_combo.addItems(["1", "2", "3", "4"])
        self.copies_combo.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border: 1px solid #E5E5EA;
                border-radius: 8px;
                font-size: 14px;
            }
        """)
        options_layout.addWidget(self.copies_combo)
        
        self.remember_path_cb = QCheckBox("记住上次选择的路径")
        self.remember_path_cb.setChecked(self.config.get('remember_path', True))
        self.remember_path_cb.setStyleSheet("""
            QCheckBox {
                font-size: 14px;
                padding: 5px;
            }
        """)
        options_layout.addWidget(self.remember_path_cb)
        
        left_layout.addWidget(options_frame)
        
        # 打印按钮和进度条
        self.print_btn = iOSStyleButton("开始打印")
        self.print_btn.setEnabled(False)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #E5E5EA;
                border-radius: 8px;
                text-align: center;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #007AFF;
                border-radius: 8px;
            }
        """)
        
        left_layout.addWidget(self.print_btn)
        left_layout.addWidget(self.progress_bar)
        
        # 日志显示
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(100)
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: white;
                border: 1px solid #E5E5EA;
                border-radius: 8px;
                padding: 8px;
                font-size: 12px;
            }
        """)
        left_layout.addWidget(QLabel("操作日志:"))
        left_layout.addWidget(self.log_text)
        
        # 右侧：文件列表（占5份）
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        self.file_list = iOSStyleListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        right_layout.addWidget(QLabel("已选择的发票文件:"))
        right_layout.addWidget(self.file_list)
        
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([200, 800])  # 1:5 的比例
        
        main_layout.addWidget(splitter)
        
        # 连接信号
        self.add_files_btn.clicked.connect(self.add_files)
        self.add_folder_btn.clicked.connect(self.add_folder)
        self.remove_selected_btn.clicked.connect(self.remove_selected)
        self.clear_all_btn.clicked.connect(self.clear_all)
        self.exit_btn.clicked.connect(self.close)  # 退出按钮连接关闭事件
        self.print_btn.clicked.connect(self.start_printing)
        self.file_list.itemSelectionChanged.connect(self.update_button_states)
        
        # 加载配置到UI
        self.load_config_to_ui()
        
        self.update_button_states()
        
    def load_config(self) -> dict:
        """加载配置文件"""
        try:
            if os.path.exists(CONFIG_LOG_FILE):
                with open(CONFIG_LOG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except:
            pass
        return {'last_path': '', 'remember_path': True, 'page_layout': '1x2', 'copies_per_invoice': '2'}
    
    def save_config(self):
        """保存配置文件"""
        self.config['remember_path'] = self.remember_path_cb.isChecked()
        # 从UI获取当前选择的布局
        current_layout_text = self.page_layout_combo.currentText()
        layout_map = {
            "1×1 (每页1张)": "1x1",
            "1×2 (每页2张，纵向上下)": "1x2",
            "1×3 (每页3张，纵向上中下)": "1x3",
            "2×2 (每页4张，2行2列)": "2x2",
            "2×3 (每页6张，2行3列)": "2x3",
            "2×4 (每页8张，2行4列)": "2x4"
        }
        self.config['page_layout'] = layout_map.get(current_layout_text, "1x2")
        self.config['copies_per_invoice'] = self.copies_combo.currentText()
        try:
            with open(CONFIG_LOG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    def load_config_to_ui(self):
        """将配置加载到UI"""
        # 页面布局
        layout_reverse_map = {
            '1x1': "1×1 (每页1张)",
            '1x2': "1×2 (每页2张，纵向上下)", 
            '1x3': "1×3 (每页3张，纵向上中下)",
            '2x2': "2×2 (每页4张，2行2列)",
            '2x3': "2×3 (每页6张，2行3列)",
            '2x4': "2×4 (每页8张，2行4列)"
        }
        layout_text = layout_reverse_map.get(self.config.get('page_layout', '1x2'), "1×2 (每页2张，纵向上下)")
        index = self.page_layout_combo.findText(layout_text)
        if index >= 0:
            self.page_layout_combo.setCurrentIndex(index)
        
        # 打印份数
        copies = self.config.get('copies_per_invoice', '2')
        if copies in ['1', '2', '3', '4']:
            self.copies_combo.setCurrentText(copies)
    
    def load_print_log(self) -> Set[str]:
        """加载打印记录（基于文件路径）"""
        try:
            if os.path.exists(PRINT_LOG_FILE):
                with open(PRINT_LOG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return set(data.get('printed_files', []))
        except:
            pass
        return set()
    
    def save_print_log(self):
        """保存打印记录（基于文件路径）"""
        try:
            data = {'printed_files': list(self.printed_invoices)}
            with open(PRINT_LOG_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    def get_available_printers(self) -> List[str]:
        """获取所有可用的打印机列表（包括网络打印机）"""
        printers = []
        try:
            if platform.system() == "Windows":
                import win32print
                # 枚举所有打印机（包括网络打印机）
                printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_CONNECTIONS | win32print.PRINTER_ENUM_LOCAL)]
            elif platform.system() == "Darwin":  # macOS
                result = subprocess.run(['lpstat', '-p'], capture_output=True, text=True)
                if result.returncode == 0:
                    lines = result.stdout.splitlines()
                    for line in lines:
                        if line.startswith('printer '):
                            printer_name = line.split()[1]
                            printers.append(printer_name)
            else:  # Linux
                result = subprocess.run(['lpstat', '-p'], capture_output=True, text=True)
                if result.returncode == 0:
                    lines = result.stdout.splitlines()
                    for line in lines:
                        if line.startswith('printer '):
                            printer_name = line.split()[1]
                            printers.append(printer_name)
        except Exception as e:
            print(f"获取打印机列表失败: {e}")
            # 如果获取失败，返回空列表，后续会使用默认打印机
            pass
        
        return printers if printers else ["默认打印机"]
    
    def get_default_printer(self) -> str:
        """获取系统默认打印机"""
        try:
            if platform.system() == "Windows":
                import win32print
                return win32print.GetDefaultPrinter()
            elif platform.system() == "Darwin":  # macOS
                result = subprocess.run(['lpstat', '-d'], capture_output=True, text=True)
                if result.returncode == 0:
                    lines = result.stdout.splitlines()
                    for line in lines:
                        if line.startswith('system default destination:'):
                            return line.split(':')[-1].strip()
                return ""
            else:  # Linux
                result = subprocess.run(['lpstat', '-d'], capture_output=True, text=True)
                if result.returncode == 0:
                    lines = result.stdout.splitlines()
                    for line in lines:
                        if line.startswith('system default destination:'):
                            return line.split(':')[-1].strip()
                return ""
        except Exception as e:
            print(f"获取默认打印机失败: {e}")
            return ""
    
    def add_files(self):
        """添加发票文件"""
        last_path = self.config.get('last_path', '') if self.remember_path_cb.isChecked() else ''
        
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择发票文件", last_path,
            "支持的文件 (*.pdf *.ofd *.png *.jpg *.jpeg *.bmp);;所有文件 (*)"
        )
        
        if files:
            if self.remember_path_cb.isChecked():
                self.config['last_path'] = os.path.dirname(files[0])
                self.save_config()
            
            self.add_invoice_files(files)
    
    def add_folder(self):
        """从文件夹导入"""
        last_path = self.config.get('last_path', '') if self.remember_path_cb.isChecked() else ''
        
        folder = QFileDialog.getExistingDirectory(
            self, "选择发票文件夹", last_path
        )
        
        if folder:
            if self.remember_path_cb.isChecked():
                self.config['last_path'] = folder
                self.save_config()
            
            # 支持的文件扩展名
            supported_ext = ('.pdf', '.ofd', '.png', '.jpg', '.jpeg', '.bmp')
            files = []
            
            for file in os.listdir(folder):
                if file.lower().endswith(supported_ext):
                    files.append(os.path.join(folder, file))
            
            self.add_invoice_files(files)
    
    def add_invoice_files(self, files: List[str]):
        """添加发票文件到列表"""
        added_count = 0
        duplicate_count = 0
        duplicate_files = []
        
        for file_path in files:
            if file_path not in self.invoice_files:
                # 检查是否重复打印（基于文件路径）
                if file_path in self.printed_invoices:
                    duplicate_files.append(file_path)
                else:
                    self.invoice_files.append(file_path)
                    # 显示文件名（清理乱码）
                    clean_name = self.get_clean_filename(os.path.basename(file_path))
                    self.file_list.addItem(clean_name)
                    added_count += 1
            else:
                duplicate_count += 1
        
        # 处理重复文件
        if duplicate_files:
            dialog = DuplicateInvoiceDialog(duplicate_files, self)
            result = dialog.exec_()
            
            if result == QDialog.Accepted:
                if dialog.result is True:  # 全部是
                    for file_path in duplicate_files:
                        self.invoice_files.append(file_path)
                        clean_name = self.get_clean_filename(os.path.basename(file_path))
                        self.file_list.addItem(clean_name)
                        added_count += 1
                elif dialog.result is False:  # 全否
                    duplicate_count += len(duplicate_files)
            else:  # 取消
                # 不添加任何重复文件
                duplicate_count += len(duplicate_files)
        
        if added_count > 0:
            self.log_message(f"添加了 {added_count} 个发票文件")
        if duplicate_count > 0:
            self.log_message(f"跳过了 {duplicate_count} 个重复文件")
        
        self.update_button_states()
    
    def get_clean_filename(self, filename: str) -> str:
        """获取干净的文件名，去除乱码"""
        # 只保留中文、数字、字母和常见符号
        clean_name = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9._()-]', '', filename)
        if not clean_name:
            # 如果清理后为空，返回原始文件名
            return filename
        return clean_name
    
    def remove_selected(self):
        """移除选中的文件"""
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            return
        
        for item in selected_items:
            row = self.file_list.row(item)
            self.invoice_files.pop(row)
            self.file_list.takeItem(row)
        
        self.log_message(f"移除了 {len(selected_items)} 个文件")
        self.update_button_states()
    
    def clear_all(self):
        """清空所有文件"""
        if self.invoice_files:
            reply = QMessageBox.question(
                self, "确认清空",
                "确定要清空所有发票文件吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.invoice_files.clear()
                self.file_list.clear()
                self.log_message("已清空所有文件")
                self.update_button_states()
    
    def update_button_states(self):
        """更新按钮状态"""
        has_files = len(self.invoice_files) > 0
        has_selection = len(self.file_list.selectedItems()) > 0
        
        self.print_btn.setEnabled(has_files)
        self.remove_selected_btn.setEnabled(has_selection)
        self.clear_all_btn.setEnabled(has_files)
    
    def log_message(self, message: str):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def start_printing(self):
        """开始打印"""
        if not self.invoice_files:
            return
        
        # 保存当前设置到配置文件
        self.save_config()
        
        if self.print_mode_combo.currentText() == "打印为PDF":
            # 获取输出路径
            output_file, _ = QFileDialog.getSaveFileName(
                self, "保存PDF文件", "",
                "PDF文件 (*.pdf)"
            )
            if not output_file:
                return
            if not output_file.endswith('.pdf'):
                output_file += '.pdf'
            
            # 启动打印线程（PDF模式）
            current_layout_text = self.page_layout_combo.currentText()
            layout_map = {
                "1×1 (每页1张)": "1x1",
                "1×2 (每页2张，纵向上下)": "1x2",
                "1×3 (每页3张，纵向上中下)": "1x3",
                "2×2 (每页4张，2行2列)": "2x2",
                "2×3 (每页6张，2行3列)": "2x3",
                "2×4 (每页8张，2行4列)": "2x4"
            }
            page_layout = layout_map.get(current_layout_text, "1x2")
            copies_per_invoice = int(self.copies_combo.currentText())
            self.print_worker = PrintWorker(self.invoice_files, output_file, True, None, page_layout, copies_per_invoice, self)
            self.setup_print_worker()
            self.print_worker.start()
            
        else:
            # 直接打印模式
            # 获取所有可用的打印机（包括网络打印机）
            available_printers = self.get_available_printers()
            
            # 获取系统默认打印机
            default_printer = self.get_default_printer()
            
            selected_printer = None
            
            # 如果有多个打印机，显示选择对话框
            if len(available_printers) > 1:
                # 显示选择对话框，默认选中系统默认打印机
                printer_dialog = PrinterSelectionDialog(available_printers, default_printer, self)
                result = printer_dialog.exec_()
                if result == QDialog.Accepted:
                    selected_printer = printer_dialog.selected_printer
                else:
                    # 用户取消了选择
                    return
            elif len(available_printers) == 1:
                # 只有一个打印机，直接使用
                selected_printer = available_printers[0] if available_printers[0] != "默认打印机" else None
            else:
                # 没有检测到打印机，使用默认打印机
                selected_printer = None
            
            # 禁用UI
            self.print_btn.setEnabled(False)
            self.add_files_btn.setEnabled(False)
            self.add_folder_btn.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # 启动打印线程（直接打印模式）
            current_layout_text = self.page_layout_combo.currentText()
            layout_map = {
                "1×1 (每页1张)": "1x1",
                "1×2 (每页2张，纵向上下)": "1x2",
                "1×3 (每页3张，纵向上中下)": "1x3",
                "2×2 (每页4张，2行2列)": "2x2",
                "2×3 (每页6张，2行3列)": "2x3",
                "2×4 (每页8张，2行4列)": "2x4"
            }
            page_layout = layout_map.get(current_layout_text, "1x2")
            copies_per_invoice = int(self.copies_combo.currentText())
            self.print_worker = PrintWorker(self.invoice_files, "", False, selected_printer, page_layout, copies_per_invoice, self)
            self.setup_print_worker()
            self.print_worker.start()
    
    def setup_print_worker(self):
        """设置打印工作线程"""
        self.print_worker.progress_updated.connect(self.update_progress)
        self.print_worker.finished.connect(self.printing_finished)
    
    def update_progress(self, value: int, message: str):
        """更新进度"""
        self.progress_bar.setValue(value)
        self.log_message(message)
    
    def printing_finished(self, success: bool, message: str, printed_count: int):
        """打印完成"""
        # 更新打印记录（基于文件路径）
        if success:
            for file_path in self.invoice_files:
                self.printed_invoices.add(file_path)
            self.save_print_log()
        
        # 恢复UI
        self.print_btn.setEnabled(True)
        self.add_files_btn.setEnabled(True)
        self.add_folder_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # 显示结果
        if success:
            current_layout_text = self.page_layout_combo.currentText()
            copies_per_invoice = self.copies_combo.currentText()
            final_message = f"{message}\n\n已处理 {printed_count} 张发票\n页面布局: {current_layout_text}\n每张发票打印 {copies_per_invoice} 份"
            QMessageBox.information(self, "完成", final_message)
            self.log_message(f"打印任务完成，共处理 {printed_count} 张发票")
        else:
            error_message = f"{message}\n\n可能的原因：\n1. 打印机未连接或未就绪\n2. 系统打印服务未启动\n3. 请检查打印机状态后重试\n\n临时文件已保存，您可以手动打印。"
            QMessageBox.critical(self, "打印失败", error_message)
            self.log_message(f"打印失败: {message}")

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # 使用Fusion风格作为基础
    
    window = InvoicePrinterApp()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()