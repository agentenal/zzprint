import sys
import os
import json
import shutil
import re
import datetime
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QComboBox, QListWidget, 
                             QLabel, QTextEdit, QFileDialog, QFrame, QSpinBox, 
                             QMessageBox, QScrollArea, QAbstractItemView,
                             QTableWidget, QTableWidgetItem, QHeaderView, QLineEdit)
from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtGui import QImage, QPixmap, QColor, QKeySequence, QFont

# --- 核心处理引擎 ---
class PrintingEngine:
    def __init__(self, ledger_file="invoice_ledger.json"):
        self.ledger_file = ledger_file
        self.ledger = self.load_ledger()
        self.layout_map = {
            "1×1": (1, 1), "1×2": (2, 1), "1×3": (3, 1),
            "2×2": (2, 2), "2×3": (3, 2), "2×4": (4, 2)
        }

    def load_ledger(self):
        if os.path.exists(self.ledger_file):
            try:
                with open(self.ledger_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data
            except: return {}
        return {}

    def save_ledger(self, info):
        if info["发票号码"] != "未知":
            # 记录打印日期，格式 yyyy/mm/dd HH:MM
            info["打印日期"] = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
            self.ledger[info["发票号码"]] = info
            with open(self.ledger_file, 'w', encoding='utf-8') as f:
                json.dump(self.ledger, f, ensure_ascii=False, indent=4)

    def normalize_date(self, date_str):
        """ 将各种日期格式统一为 yyyy/mm/dd """
        if not date_str or date_str == "未知": return "未知"
        try:
            # 清理中文和符号
            clean_str = date_str.replace("年", "/").replace("月", "/").replace("日", "").strip()
            # 尝试解析
            dt = datetime.datetime.strptime(clean_str, "%Y/%m/%d")
            return dt.strftime("%Y/%m/%d")
        except:
            return date_str

    def parse_invoice(self, file_path):
        base_info = {
            "发票号码": "未知", "开票日期": "未知", "自产农产品销售": "否",
            "购买方名称": "未知", "购买方税号": "未知",
            "销售方名称": "未知", "销售方税号": "未知",
            "备注": "无", "文件名": os.path.basename(file_path),
            "打印日期": "待处理",
            "items": []
        }
        try:
            with pdfplumber.open(file_path) as pdf:
                page = pdf.pages[0]
                text = page.extract_text()
                lines = text.split('\n')
                
                # 1. 基础信息正则提取
                m_no = re.search(r'发票号码[:：]\s*(\d+)', text)
                if m_no: base_info["发票号码"] = m_no.group(1)
                
                m_date = re.search(r'开票日期[:：]\s*(\d{4}年\d{1,2}月\d{1,2}日)', text)
                if m_date: base_info["开票日期"] = self.normalize_date(m_date.group(1))
                
                # 特殊标记识别
                if "自产农产品" in text or "自产农产品销售" in text: 
                    base_info["自产农产品销售"] = "是"

                names = re.findall(r'名称[:：]\s*([^\n\s]+)', text)
                ids = re.findall(r'纳税人识别号[:：]\s*([A-Z0-9]+)', text)
                if len(names) >= 2: base_info["购买方名称"], base_info["销售方名称"] = names[0], names[1]
                if len(ids) >= 2: base_info["购买方税号"], base_info["销售方税号"] = ids[0], ids[1]

                # 2. 明细行提取 (增强版)
                for line in lines:
                    if any(c.isdigit() for c in line) and ('*' in line or '免税' in line or '.' in line):
                        parts = line.split()
                        if len(parts) < 5: continue 
                        
                        try:
                            # 倒序取值策略
                            tax_amt_str = parts[-1] 
                            tax_rate_str = parts[-2]
                            amt_str = parts[-3]
                            
                            tax_amt = 0.00
                            if '***' in tax_amt_str: tax_amt = 0.00
                            else: tax_amt = float(tax_amt_str.replace('￥','').replace(',',''))
                            
                            amt = float(amt_str.replace('￥','').replace(',',''))
                            
                            qty = "1"
                            unit = "无"
                            price = str(amt)
                            
                            if len(parts) >= 6:
                                try:
                                    if self.is_float(parts[-4]) and self.is_float(parts[-5]):
                                        price = parts[-4]
                                        qty = parts[-5]
                                        unit = parts[-6] if not self.is_float(parts[-6]) else "无"
                                except: pass

                            base_info["items"].append({
                                "项目名称": parts[0],
                                "规格型号": "无",
                                "单位": unit,
                                "数量": qty, 
                                "单价": price,
                                "金额": f"{amt:.2f}", 
                                "税率": tax_rate_str,
                                "税额": f"{tax_amt:.2f}", 
                                "合计": f"{(amt + tax_amt):.2f}"
                            })
                        except: 
                            continue
                
                if not base_info["items"]:
                    total_m = re.search(r'[（\(]小写[）\)]\s*[￥¥]?\s*([\d\.]+)', text)
                    if total_m:
                        val = float(total_m.group(1))
                        base_info["items"].append({
                            "项目名称": "（总额识别）", "数量": "1", "单价": val, 
                            "金额": f"{val:.2f}", "税率": "-", "税额": "0.00", "合计": f"{val:.2f}"
                        })
        except: pass
        return base_info

    def is_float(self, s):
        try: float(s.replace(',','')); return True
        except: return False

    def create_layout(self, input_files, layout_desc, output_path, copies=1):
        a4_w, a4_h = 595, 842 
        doc = fitz.open()
        rows, cols = self.layout_map.get(layout_desc, (1, 1))
        expanded = [f for f in input_files for _ in range(copies)]
        for i in range(0, len(expanded), rows * cols):
            page = doc.new_page(width=a4_w, height=a4_h)
            batch = expanded[i:i + rows * cols]
            cw, ch = a4_w / cols, a4_h / rows
            for idx, f_path in enumerate(batch):
                try:
                    src = fitz.open(f_path)
                    r, c = divmod(idx, cols)
                    rect = fitz.Rect(c*cw+10, r*ch+10, (c+1)*cw-10, (r+1)*ch-10)
                    page.show_pdf_page(rect, src, 0)
                    src.close()
                except: pass
        doc.save(output_path); doc.close()

# --- 增强型表格控件（支持复制） ---
class CopyableTable(QTableWidget):
    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Copy):
            indices = self.selectedIndexes()
            if not indices: return
            rows = sorted(list(set(i.row() for i in indices)))
            cols = sorted(list(set(i.column() for i in indices)))
            table_text = ""
            for r in rows:
                row_data = []
                for c in cols:
                    item = self.item(r, c)
                    row_data.append(item.text() if item else "")
                table_text += "\t".join(row_data) + "\n"
            QApplication.clipboard().setText(table_text)
        else: super().keyPressEvent(event)

# --- 主程序界面 ---
class ZZPrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.engine = PrintingEngine()
        self.settings = QSettings("ZZStudio", "ZZPrinterV2")
        self.setWindowTitle("ZZ打票助手 - 专业增强版")
        self.setMinimumSize(1280, 850)
        self.setAcceptDrops(True)
        
        # 状态变量
        self.group_stat_active = False
        self.summary_level = 1
        self.theme_mode = self.settings.value("theme", "light") 
        
        # 排序状态
        self.sort_col = "打印日期" 
        self.sort_asc = False 

        self.init_ui()
        self.apply_theme() 
        self.refresh_table()

    def init_ui(self):
        central_widget = QWidget(); self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # === 左侧控制面板 ===
        side_scroll = QScrollArea()
        side_scroll.setFixedWidth(300)
        side_scroll.setWidgetResizable(True)
        side_scroll.setFrameShape(QFrame.Shape.NoFrame)
        
        side_content = QFrame(); side_content.setObjectName("SidePanel")
        side_layout = QVBoxLayout(side_content)
        side_layout.setSpacing(12)
        side_layout.setContentsMargins(15, 20, 15, 20)
        
        # 标题栏
        header_layout = QHBoxLayout()
        title_label = QLabel("ZZ 打票助手")
        title_label.setFont(QFont("Microsoft YaHei", 18, QFont.Weight.Bold))
        header_layout.addWidget(title_label)
        self.btn_theme = QPushButton("🌓"); self.btn_theme.setFixedWidth(40)
        self.btn_theme.clicked.connect(self.toggle_theme); header_layout.addWidget(self.btn_theme)
        side_layout.addLayout(header_layout)
        
        # 文件操作区
        side_layout.addWidget(QLabel("📂 文件管理"))
        btn_add_file = self.create_btn("添加发票文件", self.add_files, "#007AFF")
        btn_add_dir = self.create_btn("从文件夹导入", self.add_folder, "#5856D6")
        btn_rm_sel = self.create_btn("移除选中文件", self.remove_selected, "#FF9500")
        
        self.btn_remove_dup = self.create_btn("一键移除已打印", self.remove_duplicates, "#FF3B30")
        self.btn_remove_dup.setEnabled(False)
        
        side_layout.addWidget(btn_add_file)
        side_layout.addWidget(btn_add_dir)
        side_layout.addWidget(btn_rm_sel)
        side_layout.addWidget(self.btn_remove_dup)

        side_layout.addWidget(self.create_line())

        # 打印设置区
        side_layout.addWidget(QLabel("🖨️ 打印设置"))
        self.mode_combo = QComboBox(); self.mode_combo.addItems(["直接打印", "打印为PDF"])
        self.mode_combo.setFixedHeight(35)
        side_layout.addWidget(QLabel("输出模式:"))
        side_layout.addWidget(self.mode_combo)
        
        self.layout_combo = QComboBox(); self.layout_combo.addItems(["1×1", "1×2", "1×3", "2×2", "2×3", "2×4"])
        self.layout_combo.setFixedHeight(35)
        self.layout_combo.currentTextChanged.connect(self.update_preview)
        side_layout.addWidget(QLabel("页面布局:"))
        side_layout.addWidget(self.layout_combo)
        
        self.copy_spin = QSpinBox(); self.copy_spin.setRange(1, 10); self.copy_spin.setValue(1)
        self.copy_spin.setFixedHeight(35)
        side_layout.addWidget(QLabel("单张份数:"))
        side_layout.addWidget(self.copy_spin)

        side_layout.addWidget(self.create_line())

        # 数据操作区
        side_layout.addWidget(QLabel("📊 数据台账"))
        btn_excel = self.create_btn("导出 Excel 台账", self.export_excel, "#34C759")
        btn_clear = self.create_btn("清空打印队列", self.clear_all, "#8E8E93")
        side_layout.addWidget(btn_excel)
        side_layout.addWidget(btn_clear)

        side_layout.addStretch()
        
        # 底部操作
        self.btn_print = QPushButton("🚀 开始处理 / 打印")
        self.btn_print.setFixedHeight(50)
        self.btn_print.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
        self.btn_print.clicked.connect(self.process_printing)
        side_layout.addWidget(self.btn_print)
        
        btn_quit = QPushButton("退出程序"); btn_quit.setObjectName("QuitBtn")
        btn_quit.setFixedHeight(35)
        btn_quit.clicked.connect(self.close); side_layout.addWidget(btn_quit)

        side_scroll.setWidget(side_content); main_layout.addWidget(side_scroll)

        # === 右侧内容区 ===
        content_layout = QVBoxLayout()
        
        top_split = QHBoxLayout()
        self.file_list = QListWidget(); self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        v1 = QVBoxLayout(); v1.addWidget(QLabel("<b>📄 待打印队列</b>")); v1.addWidget(self.file_list); top_split.addLayout(v1, 4)
        
        self.preview_label = QLabel("预览区"); self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scroll_area = QScrollArea(); self.scroll_area.setWidget(self.preview_label); self.scroll_area.setWidgetResizable(True)
        v2 = QVBoxLayout(); v2.addWidget(QLabel("<b>👁️ 实时预览</b>")); v2.addWidget(self.scroll_area); top_split.addLayout(v2, 3)
        content_layout.addLayout(top_split, 3)

        # 筛选栏
        filter_box = QFrame(); filter_box.setObjectName("FilterBox")
        filter_box.setFixedHeight(50)
        filter_layout = QHBoxLayout(filter_box)
        filter_layout.setContentsMargins(10, 5, 10, 5)

        self.search_seller = QLineEdit(); self.search_seller.setPlaceholderText("🔍 筛选销售方...")
        self.search_seller.textChanged.connect(self.refresh_table); filter_layout.addWidget(self.search_seller, 2)
        
        self.search_date = QLineEdit(); self.search_date.setPlaceholderText("📅 开票年月(202601)")
        self.search_date.textChanged.connect(self.refresh_table); filter_layout.addWidget(self.search_date, 1)

        filter_layout.addWidget(QLabel("|"))

        self.btn_group_stat = QPushButton("📊 分组统计: 关"); self.btn_group_stat.setCheckable(True)
        self.btn_group_stat.clicked.connect(self.toggle_group_stat); filter_layout.addWidget(self.btn_group_stat)
        
        self.btn_sum_level = QPushButton("📑 汇总: 一级"); self.btn_sum_level.clicked.connect(self.toggle_sum_level)
        self.btn_sum_level.setEnabled(False) 
        filter_layout.addWidget(self.btn_sum_level)
        
        btn_reset = QPushButton("重置条件"); btn_reset.clicked.connect(self.reset_filters)
        filter_layout.addWidget(btn_reset)
        
        content_layout.addWidget(filter_box)

        # 表格
        self.table = CopyableTable()
        self.cols = ["发票号码", "开票日期", "打印日期", "销售方名称", "自产农产品", "项目名称", "数量", "单价", "金额", "税率", "税额", "价税合计"]
        self.table.setColumnCount(len(self.cols)); self.table.setHorizontalHeaderLabels(self.cols)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().sectionClicked.connect(self.handle_header_click) 
        
        self.table.setColumnWidth(0, 120) 
        self.table.setColumnWidth(3, 180) 
        content_layout.addWidget(self.table, 4)
        
        main_layout.addLayout(content_layout)

    def create_btn(self, text, func, color_hex):
        btn = QPushButton(text)
        btn.clicked.connect(func)
        btn.setProperty("base_color", color_hex)
        return btn

    def create_line(self):
        line = QFrame(); line.setFrameShape(QFrame.Shape.HLine); line.setFrameShadow(QFrame.Shadow.Sunken)
        return line

    # --- 逻辑控制 ---

    def toggle_theme(self):
        self.theme_mode = "light" if self.theme_mode == "dark" else "dark"
        self.settings.setValue("theme", self.theme_mode)
        self.apply_theme()

    def toggle_group_stat(self):
        self.group_stat_active = self.btn_group_stat.isChecked()
        self.btn_group_stat.setText("📊 分组统计: 开" if self.group_stat_active else "📊 分组统计: 关")
        self.btn_sum_level.setEnabled(self.group_stat_active)
        self.refresh_table()

    def toggle_sum_level(self):
        self.summary_level = 2 if self.summary_level == 1 else 1
        self.btn_sum_level.setText("📑 汇总: 二级" if self.summary_level == 2 else "📑 汇总: 一级")
        if self.group_stat_active: self.refresh_table()

    def handle_header_click(self, index):
        if self.group_stat_active:
            QMessageBox.information(self, "提示", "分组模式下按销售方+日期固定排序。")
            return

        col_name = self.cols[index]
        if self.sort_col == col_name:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_col = col_name
            self.sort_asc = True 
            
        for i in range(self.table.columnCount()):
            item = self.table.horizontalHeaderItem(i)
            txt = self.cols[i]
            if txt == self.sort_col:
                item.setText(f"{txt} {'↑' if self.sort_asc else '↓'}")
            else:
                item.setText(txt)
        
        self.refresh_table()

    def get_data_frame(self):
        seller_key = self.search_seller.text().strip().lower()
        date_key = self.search_date.text().strip().replace("-", "").replace("年","").replace("月","")
        
        raw_list = []
        for no, base in self.engine.ledger.items():
            s_name = base.get("销售方名称", "").lower()
            d_date = base.get("开票日期", "").replace("/", "").replace("-","")
            
            if seller_key in s_name and date_key in d_date:
                for item in base.get("items", []):
                    row = base.copy()
                    del row["items"] 
                    row.update(item) 
                    try: row["金额"] = float(row["金额"])
                    except: row["金额"] = 0.0
                    try: row["税额"] = float(row["税额"])
                    except: row["税额"] = 0.0
                    try: row["合计"] = float(row["合计"])
                    except: row["合计"] = 0.0
                    
                    row["价税合计"] = row["合计"]
                    raw_list.append(row)
        
        return pd.DataFrame(raw_list)

    def refresh_table(self):
        self.table.setRowCount(0)
        df = self.get_data_frame()
        if df.empty: return

        if not self.group_stat_active:
            if self.sort_col in df.columns:
                df = df.sort_values(by=self.sort_col, ascending=self.sort_asc)
            self._fill_rows_from_df(df)
            
        else:
            bg_l1 = QColor("#E3F2FD") if self.theme_mode == "light" else QColor("#0D47A1") # 销售方小计
            bg_l2 = QColor("#F1F8E9") if self.theme_mode == "light" else QColor("#1B5E20") # 日期小计
            
            # 按销售方（字母顺序） + 打印日期（倒序，最近的日期在前）进行排序
            df = df.sort_values(by=["销售方名称", "打印日期"], ascending=[True, False])
            
            grouped_seller = df.groupby("销售方名称", sort=False)
            
            for seller_name, group_df in grouped_seller:
                
                # --- 一级汇总 ---
                if self.summary_level == 1:
                    self._fill_rows_from_df(group_df)
                    self._insert_sum_row(f"【{seller_name}】 总计", group_df, bg_l1)
                
                # --- 二级汇总 (按打印日期) ---
                else:
                    # 创建一个临时列用于按“天”分组（忽略时分秒）
                    # 使用 .copy() 防止 SettingWithCopyWarning
                    group_df_c = group_df.copy()
                    group_df_c["_day_group"] = group_df_c["打印日期"].apply(lambda x: str(x).split(" ")[0])
                    
                    # 按天分组
                    grouped_date = group_df_c.groupby("_day_group", sort=False)
                    
                    for date_val, date_df in grouped_date:
                        self._fill_rows_from_df(date_df)
                        # 插入日期小计
                        self._insert_sum_row(f"  └─ 日期 {date_val} 小计", date_df, bg_l2)
                    
                    # 最后插入销售方总计
                    self._insert_sum_row(f"【{seller_name}】 总计", group_df, bg_l1)

    def _fill_rows_from_df(self, df):
        for _, row_data in df.iterrows():
            r = self.table.rowCount()
            self.table.insertRow(r)
            for i, col_key in enumerate(self.cols):
                val = row_data.get(col_key, "")
                if isinstance(val, float): val = f"{val:.2f}"
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(r, i, item)

    def _insert_sum_row(self, label, df_scope, bg_color):
        r = self.table.rowCount()
        self.table.insertRow(r)
        
        label_item = QTableWidgetItem(label)
        label_item.setBackground(bg_color)
        label_item.setFont(QFont("Microsoft YaHei", 9, QFont.Weight.Bold))
        self.table.setItem(r, 3, label_item) 
        
        sum_cols = {"金额": 8, "税额": 10, "价税合计": 11}
        for col_name, col_idx in sum_cols.items():
            val = df_scope[col_name].sum()
            item = QTableWidgetItem(f"{val:.2f}")
            item.setBackground(bg_color)
            item.setForeground(QColor("#D32F2F") if self.theme_mode == "light" else QColor("#FF6659"))
            item.setFont(QFont("Microsoft YaHei", 9, QFont.Weight.Bold))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(r, col_idx, item)
            
        for i in range(self.table.columnCount()):
            if self.table.item(r, i) is None:
                empty = QTableWidgetItem("")
                empty.setBackground(bg_color)
                self.table.setItem(r, i, empty)

    def export_excel(self):
        df = self.get_data_frame()
        if df.empty:
            QMessageBox.warning(self, "提示", "没有数据可导出！")
            return
        
        p, _ = QFileDialog.getSaveFileName(self, "保存台账", f"发票台账_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx", "*.xlsx")
        if p:
            try:
                export_cols = ["发票号码", "打印日期", "开票日期", "销售方名称", "销售方税号", "购买方名称", 
                               "自产农产品销售", "项目名称", "规格型号", "单位", "数量", "单价", 
                               "金额", "税率", "税额", "价税合计", "备注", "文件名"]
                for c in export_cols:
                    if c not in df.columns: df[c] = ""
                df[export_cols].to_excel(p, index=False)
                QMessageBox.information(self, "成功", "导出成功！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")

    def process_printing(self):
        paths = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if not paths: return
        
        mode = self.mode_combo.currentText()
        out_file = "temp_print_task.pdf"
        
        self.engine.create_layout(paths, self.layout_combo.currentText(), out_file, self.copy_spin.value())
        
        if mode == "打印为PDF":
            save_p, _ = QFileDialog.getSaveFileName(self, "保存打印文件", f"打印任务_{datetime.datetime.now().strftime('%H%M%S')}.pdf", "*.pdf")
            if save_p:
                shutil.move(out_file, save_p)
                self._mark_as_printed()
                os.startfile(save_p)
        else:
            QMessageBox.information(self, "提示", "已生成打印文件，将打开预览，请在打开的窗口中点击打印。")
            self._mark_as_printed()
            os.startfile(out_file)
            
    def _mark_as_printed(self):
        for i in range(self.file_list.count()): 
            info = self.file_list.item(i).data(Qt.ItemDataRole.UserRole)
            self.engine.save_ledger(info)
        self.refresh_table()

    def remove_duplicates(self):
        for i in range(self.file_list.count() - 1, -1, -1):
            if self.file_list.item(i).data(Qt.ItemDataRole.UserRole)["发票号码"] in self.engine.ledger:
                self.file_list.takeItem(i)
        self.btn_remove_dup.setEnabled(False)
        self.update_preview()

    def handle_files(self, paths):
        has_dup = False
        for p in paths:
            if p.lower().endswith(('.pdf', '.ofd')):
                info = self.engine.parse_invoice(p)
                list_item = QListWidget() 
                self.file_list.addItem(p)
                row = self.file_list.count() - 1
                self.file_list.item(row).setData(Qt.ItemDataRole.UserRole, info)
                
                if info["发票号码"] in self.engine.ledger:
                    self.file_list.item(row).setForeground(QColor("#FF3B30"))
                    self.file_list.item(row).setToolTip("该发票已打印过！")
                    has_dup = True
                else:
                    self.file_list.item(row).setToolTip(f"发票号: {info['发票号码']}")

        if has_dup: self.btn_remove_dup.setEnabled(True)
        self.update_preview()

    def apply_theme(self):
        is_dark = self.theme_mode == "dark"
        
        c = {
            "bg": "#1E1E1E" if is_dark else "#F5F5F7",
            "panel": "#2D2D2D" if is_dark else "#FFFFFF",
            "text": "#FFFFFF" if is_dark else "#333333",
            "border": "#404040" if is_dark else "#E5E5EA",
            "input_bg": "#3A3A3A" if is_dark else "#F2F2F7",
            "table_head": "#333333" if is_dark else "#E5E5EA",
        }

        style = f"""
            QMainWindow, QWidget {{ background-color: {c["bg"]}; color: {c["text"]}; font-family: "Microsoft YaHei", sans-serif; }}
            
            QFrame#SidePanel, QFrame#FilterBox {{ 
                background-color: {c["panel"]}; 
                border-radius: 12px; 
                border: 1px solid {c["border"]}; 
            }}
            
            QLineEdit, QComboBox, QSpinBox {{
                background-color: {c["input_bg"]};
                border: 1px solid {c["border"]};
                border-radius: 6px;
                padding: 4px 8px;
                color: {c["text"]};
                selection-background-color: #007AFF;
            }}
            
            QPushButton {{
                background-color: #007AFF; 
                color: white; 
                border-radius: 6px; 
                padding: 6px 12px; 
                border: none;
                font-weight: bold;
            }}
            QPushButton:hover {{ opacity: 0.8; }}
            QPushButton:pressed {{ opacity: 0.6; }}
            QPushButton:disabled {{ background-color: {c["border"]}; color: #999; }}
            QPushButton#QuitBtn {{ background-color: #FF3B30; }}
            
            QListWidget {{
                background-color: {c["panel"]};
                border: 1px solid {c["border"]};
                border-radius: 8px;
                outline: none;
            }}
            QListWidget::item {{ height: 28px; padding-left: 5px; }}
            QListWidget::item:selected {{ background-color: #007AFF; color: white; }}
            
            QTableWidget {{
                background-color: {c["panel"]};
                gridline-color: {c["border"]};
                border: 1px solid {c["border"]};
                border-radius: 8px;
                selection-background-color: #B3D7FF;
                selection-color: black;
            }}
            QHeaderView::section {{
                background-color: {c["table_head"]};
                padding: 5px;
                border: none;
                font-weight: bold;
                border-right: 1px solid {c["border"]};
                border-bottom: 1px solid {c["border"]};
            }}
            QScrollBar:vertical {{ width: 10px; background: transparent; }}
            QScrollBar::handle:vertical {{ background: #999; border-radius: 5px; }}
        """
        self.setStyleSheet(style)
        
        buttons = self.findChildren(QPushButton)
        for btn in buttons:
            base_color = btn.property("base_color")
            if base_color:
                btn.setStyleSheet(f"background-color: {base_color}; color: white; border-radius: 6px; padding: 6px;")

    def update_preview(self):
        if self.file_list.count() == 0: self.preview_label.setPixmap(QPixmap()); return
        paths = [self.file_list.item(i).text() for i in range(min(self.file_list.count(), 6))]
        self.engine.create_layout(paths, self.layout_combo.currentText(), "pre.pdf", self.copy_spin.value())
        try:
            doc = fitz.open("pre.pdf")
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
            self.preview_label.setPixmap(QPixmap.fromImage(img).scaledToWidth(self.scroll_area.width()-30, Qt.TransformationMode.SmoothTransformation))
            doc.close()
        except: pass

    def add_files(self): 
        f, _ = QFileDialog.getOpenFileNames(self, "选择发票", "", "PDF/OFD (*.pdf *.ofd)")
        if f: self.handle_files(f)
    def add_folder(self): 
        d = QFileDialog.getExistingDirectory(self)
        if d: self.handle_files([os.path.join(d, x) for x in os.listdir(d) if x.lower().endswith(('.pdf', '.ofd'))])
    def remove_selected(self):
        for i in self.file_list.selectedItems(): self.file_list.takeItem(self.file_list.row(i))
        self.update_preview()
    def clear_all(self): self.file_list.clear(); self.update_preview()
    def reset_filters(self): self.search_seller.clear(); self.search_date.clear(); self.refresh_table()
    def dragEnterEvent(self, e): e.accept() if e.mimeData().hasUrls() else e.ignore()
    def dropEvent(self, e): self.handle_files([u.toLocalFile() for u in e.mimeData().urls()])

if __name__ == "__main__":
    app = QApplication(sys.argv); app.setStyle("Fusion"); win = ZZPrinterApp(); win.show(); sys.exit(app.exec())