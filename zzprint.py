import sys
import os
import hashlib
import json
import shutil
import re
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QComboBox, QListWidget, 
                             QLabel, QTextEdit, QFileDialog, QFrame, QSpinBox, 
                             QMessageBox, QScrollArea, QAbstractItemView,
                             QTableWidget, QTableWidgetItem, QHeaderView, QLineEdit)
from PyQt6.QtCore import Qt, QRectF, QSettings
from PyQt6.QtGui import QImage, QPixmap, QPainter, QPalette, QColor

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
                    return json.load(f)
            except: return {}
        return {}

    def save_ledger(self, info):
        if info["发票号码"] != "未知":
            self.ledger[info["发票号码"]] = info
            with open(self.ledger_file, 'w', encoding='utf-8') as f:
                json.dump(self.ledger, f, ensure_ascii=False, indent=4)

    def parse_invoice(self, file_path):
        """ 高精度解析逻辑 (已修复金额识别) """
        info = {
            "发票号码": "未知", "开票日期": "未知", "自产农产品销售": "否",
            "购买方名称": "未知", "购买方税号": "未知",
            "销售方名称": "未知", "销售方税号": "未知",
            "项目名称": "未知", "规格型号": "无", "单位": "无",
            "数量": "0", "单价": "0", "金额": "0.00",
            "税率": "免税", "税额": "0.00", "合计": "0.00",
            "备注": "无", "开票人": "未知", "文件名": os.path.basename(file_path)
        }
        try:
            with pdfplumber.open(file_path) as pdf:
                text = pdf.pages[0].extract_text()
                lines = text.split('\n')
                
                # 1. 基础信息
                m_no = re.search(r'发票号码[:：]\s*(\d+)', text)
                if m_no: info["发票号码"] = m_no.group(1)
                
                m_date = re.search(r'开票日期[:：]\s*(\d{4}年\d{1,2}月\d{1,2}日)', text)
                if m_date: info["开票日期"] = m_date.group(1)
                
                if "自产农产品销售" in text: info["自产农产品销售"] = "是"

                # 2. 购销双方信息
                names = re.findall(r'名称[:：]\s*([^\n\s]+)', text)
                ids = re.findall(r'纳税人识别号[:：]\s*([A-Z0-9]+)', text)
                if len(names) >= 2: info["购买方名称"], info["销售方名称"] = names[0], names[1]
                if len(ids) >= 2: info["购买方税号"], info["销售方税号"] = ids[0], ids[1]

                # 3. 合计金额 (关键修复：兼容中文括号和英文括号)
                # 匹配 "(小写)" 或 "（小写）" 后面的数字
                total_m = re.search(r'[（\(]小写[）\)]\s*[￥¥]?\s*([\d\.]+)', text)
                if total_m: 
                    info["合计"] = total_m.group(1)
                else:
                    # 备用方案：如果上面没匹配到，尝试匹配 "小写" 附近的金额
                    backup_m = re.search(r'小写.*?[￥¥]?\s*([\d\.]+)', text)
                    if backup_m: info["合计"] = backup_m.group(1)

                # 4. 明细行解析
                for line in lines:
                    # 只有包含数字且有 * 号的行才可能是明细
                    if '*' in line and any(c.isdigit() for c in line):
                        parts = line.split()
                        if len(parts) >= 6:
                            try:
                                info["项目名称"] = parts[0]
                                info["税额"] = parts[-1].replace('***', '0.00')
                                info["税率"] = parts[-2]
                                info["金额"] = parts[-3]
                                info["单价"] = parts[-4]
                                info["数量"] = parts[-5]
                                if len(parts) == 7: info["单位"] = parts[1]
                                elif len(parts) >= 8:
                                    info["规格型号"] = parts[1]; info["单位"] = parts[2]
                            except: pass
                
                # 5. 备注
                bz_m = re.search(r'备\s*注\s*([\s\S]+?)\s*收\s*款\s*人', text)
                if bz_m: info["备注"] = re.sub(r'开票人[:：].*$', '', bz_m.group(1).strip().replace('\n', ' ')).strip()
        except: pass
        return info

    def create_layout(self, input_files, layout_desc, output_path, copies=1):
        a4_w, a4_h = 595, 842 
        doc = fitz.open()
        rows, cols = self.layout_map.get(layout_desc, (1, 1))
        expanded = [f for f in input_files for _ in range(copies)]
        if not expanded: return
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

# --- 主界面 ---
class ZZPrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.engine = PrintingEngine()
        self.settings = QSettings("ZZStudio", "ZZPrinter")
        self.setWindowTitle("ZZ打票 - 3.2 Pro 金额增强版")
        self.setMinimumSize(1200, 900)
        self.setAcceptDrops(True)
        
        self.init_ui()
        self.load_user_config()
        self.apply_theme() 
        self.refresh_table()

    def apply_theme(self):
        palette = QApplication.palette()
        dark = palette.color(QPalette.ColorRole.Window).lightness() < 128
        cfg = {
            "bg": "#1C1C1E" if dark else "#F2F2F7",
            "panel": "#2C2C2E" if dark else "#FFFFFF",
            "text": "#FFFFFF" if dark else "#1C1C1E",
            "border": "#3A3A3C" if dark else "#D1D1D6",
            "item_bg": "#3A3A3C" if dark else "#FFFFFF",
            "btn_sec": "#3A3A3C" if dark else "#E5E5EA",
            "log_bg": "#1C1C1E" if dark else "#F2F2F7",
            "table_head": "#48484A" if dark else "#E5E5EA",
            "input_bg": "#1C1C1E" if dark else "#E5E5EA"
        }
        self.setStyleSheet(f"""
            QMainWindow {{ background-color: {cfg['bg']}; }}
            QLabel {{ color: {cfg['text']}; font-family: "Microsoft YaHei"; }} 
            QFrame#SidePanel {{ background-color: {cfg['panel']}; border-radius: 20px; margin: 5px; }}
            QComboBox, QSpinBox, QLineEdit {{
                border: 1px solid {cfg['border']}; border-radius: 8px; padding: 6px;
                background: {cfg['item_bg']}; color: {cfg['text']}; min-height: 25px;
            }}
            QPushButton {{
                background-color: #007AFF; color: white; border-radius: 12px;
                padding: 10px; font-weight: bold; font-size: 13px; border: none;
            }}
            QPushButton:hover {{ background-color: #0062CC; }}
            QPushButton#ActionBtn {{ background-color: {cfg['btn_sec']}; color: #007AFF; }}
            QPushButton#QuitBtn {{ background-color: #FF3B30; color: white; }}
            QTableWidget {{ 
                background-color: {cfg['item_bg']}; color: {cfg['text']}; 
                gridline-color: {cfg['border']}; border: 1px solid {cfg['border']}; border-radius: 12px;
            }}
            QHeaderView::section {{
                background-color: {cfg['table_head']}; color: {cfg['text']};
                padding: 4px; border: none; font-weight: bold;
            }}
        """)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # --- 左侧面板 ---
        side_panel = QFrame(); side_panel.setObjectName("SidePanel"); side_panel.setFixedWidth(280)
        side_layout = QVBoxLayout(side_panel)
        side_layout.addWidget(QLabel("<h2 style='color:#007AFF;'>ZZ 打票 3.2</h2>"))
        
        for txt, func in [("添加发票文件", self.add_files), ("从文件夹导入", self.add_folder), ("移除选中文件", self.remove_selected)]:
            btn = QPushButton(txt); btn.clicked.connect(func); side_layout.addWidget(btn)
        
        btn_excel = QPushButton("导出 Excel 台账")
        btn_excel.setStyleSheet("background-color: #34C759; color: white;")
        btn_excel.clicked.connect(self.export_excel); side_layout.addWidget(btn_excel)

        btn_clear = QPushButton("清空列表")
        btn_clear.setObjectName("ActionBtn"); btn_clear.clicked.connect(self.clear_all); side_layout.addWidget(btn_clear)

        side_layout.addSpacing(15)
        self.mode_combo = QComboBox(); self.mode_combo.addItems(["直接打印", "打印为PDF"])
        side_layout.addWidget(QLabel("打印模式:")); side_layout.addWidget(self.mode_combo)

        self.layout_combo = QComboBox(); self.layout_combo.addItems(["1×1", "1×2", "1×3", "2×2", "2×3", "2×4"])
        self.layout_combo.currentTextChanged.connect(self.on_layout_changed)
        side_layout.addWidget(QLabel("页面布局:")); side_layout.addWidget(self.layout_combo)

        self.copy_spin = QSpinBox(); self.copy_spin.setRange(1, 4)
        self.copy_spin.valueChanged.connect(self.on_copy_changed)
        side_layout.addWidget(QLabel("单张打印份数:")); side_layout.addWidget(self.copy_spin)

        side_layout.addStretch()
        self.log_area = QTextEdit(); self.log_area.setFixedHeight(60); self.log_area.setReadOnly(True)
        side_layout.addWidget(self.log_area)

        self.btn_print = QPushButton("开始处理 / 打印"); self.btn_print.setFixedHeight(45); self.btn_print.clicked.connect(self.process_printing)
        side_layout.addWidget(self.btn_print)

        btn_quit = QPushButton("退出程序"); btn_quit.setObjectName("QuitBtn"); btn_quit.setFixedHeight(40); btn_quit.clicked.connect(self.close)
        side_layout.addWidget(btn_quit)

        # --- 右侧内容区 ---
        content_layout = QVBoxLayout()
        top_split = QHBoxLayout()
        self.file_list = QListWidget(); self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        v1 = QVBoxLayout(); v1.addWidget(QLabel("<b>发票队列</b>")); v1.addWidget(self.file_list)
        top_split.addLayout(v1, 1)

        self.preview_label = QLabel("预览区"); self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scroll_area = QScrollArea(); self.scroll_area.setWidgetResizable(True); self.scroll_area.setWidget(self.preview_label)
        v2 = QVBoxLayout(); v2.addWidget(QLabel("<b>打印预览</b>")); v2.addWidget(self.scroll_area)
        top_split.addLayout(v2, 1)
        content_layout.addLayout(top_split, 2)

        # --- 台账筛选工具栏 ---
        filter_box = QFrame(); filter_box.setStyleSheet("background: transparent; margin-top: 5px;")
        filter_layout = QHBoxLayout(filter_box); filter_layout.setContentsMargins(0,0,0,0)
        
        filter_layout.addWidget(QLabel("<b>筛选销售方:</b>"))
        self.search_seller = QLineEdit(); self.search_seller.setPlaceholderText("输入单位名称...")
        self.search_seller.textChanged.connect(self.refresh_table); filter_layout.addWidget(self.search_seller, 1)
        
        filter_layout.addWidget(QLabel("<b>筛选日期:</b>"))
        self.search_date = QLineEdit(); self.search_date.setPlaceholderText("如: 2026-01")
        self.search_date.textChanged.connect(self.refresh_table); filter_layout.addWidget(self.search_date, 1)
        
        btn_reset = QPushButton("重置"); btn_reset.setFixedWidth(60); btn_reset.setObjectName("ActionBtn")
        btn_reset.clicked.connect(self.reset_filters); filter_layout.addWidget(btn_reset)
        
        content_layout.addWidget(filter_box)

        # 台账表格
        self.table = QTableWidget()
        headers = ["发票号码", "开票日期", "销售方", "项目名称", "规格", "数量", "金额", "合计"]
        self.table.setColumnCount(len(headers)); self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        content_layout.addWidget(self.table, 1)

        main_layout.addWidget(side_panel)
        main_layout.addLayout(content_layout)

    def reset_filters(self):
        self.search_seller.clear(); self.search_date.clear(); self.refresh_table()

    def get_filtered_data(self):
        """ 获取当前筛选后的台账数据列表 """
        seller_key = self.search_seller.text().strip().lower()
        date_key = self.search_date.text().strip().replace("-", "").replace("年", "").replace("月", "")
        
        results = []
        for no, d in self.engine.ledger.items():
            s_name = d.get("销售方名称", "").lower()
            s_date = d.get("开票日期", "").replace("年", "").replace("月", "").replace("日", "")
            
            # 安全检查：防止None类型报错
            if s_name is None: s_name = ""
            if s_date is None: s_date = ""

            if seller_key in s_name and date_key in s_date:
                results.append(d)
        
        return sorted(results, key=lambda x: x.get('开票日期', ''), reverse=True)

    def refresh_table(self):
        filtered = self.get_filtered_data()
        self.table.setRowCount(0)
        for d in filtered:
            r = self.table.rowCount(); self.table.insertRow(r)
            # 这里的顺序仅用于 UI 显示，不影响导出
            vals = [d.get("发票号码"), d.get("开票日期"), d.get("销售方名称"), d.get("项目名称"), 
                    d.get("规格型号"), d.get("数量"), d.get("金额"), d.get("合计")]
            for i, v in enumerate(vals): self.table.setItem(r, i, QTableWidgetItem(str(v)))

    def export_excel(self):
        """ 导出用户指定的 16 列顺序 """
        data_list = self.get_filtered_data()
        if not data_list: 
            QMessageBox.warning(self, "提示", "当前筛选条件下没有数据可导出！")
            return

        column_order = [
            "发票号码", "开票日期", "销售方名称", "销售方税号", "自产农产品销售",
            "项目名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额", "合计",
            "购买方名称", "购买方税号"
        ]

        total_count = len(self.engine.ledger)
        filtered_count = len(data_list)
        
        msg = f"全部台账共 {total_count} 条，当前筛选结果 {filtered_count} 条。\n是否导出当前筛选后的数据？"
        reply = QMessageBox.question(self, '导出确认', msg, QMessageBox.StandardButton.Yes | 
                                     QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel)
        
        if reply == QMessageBox.StandardButton.Cancel: return
        target_data = data_list if reply == QMessageBox.StandardButton.Yes else list(self.engine.ledger.values())

        p, _ = QFileDialog.getSaveFileName(self, "保存台账", "发票台账导出.xlsx", "*.xlsx")
        if p:
            try:
                df = pd.DataFrame(target_data)
                for col in column_order:
                    if col not in df.columns: df[col] = ""
                # 强制转换为数字格式以便统计，如果需要的话
                # df["合计"] = pd.to_numeric(df["合计"], errors='ignore')
                df[column_order].to_excel(p, index=False)
                # os.startfile(os.path.dirname(p))
                self.log_area.append(f"成功导出 {len(target_data)} 条记录")
            except Exception as e:
                QMessageBox.critical(self, "导出失败", str(e))

    def update_preview(self):
        if self.file_list.count() == 0:
            self.preview_label.setText("将文件拖拽至此处添加"); self.preview_label.setPixmap(QPixmap()); return
        paths = [self.file_list.item(i).text() for i in range(min(self.file_list.count(), 6))]
        try:
            self.engine.create_layout(paths, self.layout_combo.currentText(), "pre.pdf", self.copy_spin.value())
            doc = fitz.open("pre.pdf")
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
            self.preview_label.setPixmap(QPixmap.fromImage(img).scaledToWidth(self.scroll_area.width()-40, Qt.TransformationMode.SmoothTransformation))
            doc.close()
        except: pass

    def handle_files(self, paths):
        for p in paths:
            if p.lower().endswith(('.pdf', '.ofd')):
                info = self.engine.parse_invoice(p)
                self.file_list.addItem(p)
                self.file_list.item(self.file_list.count()-1).setData(Qt.ItemDataRole.UserRole, info)
                if info["发票号码"] in self.engine.ledger: self.file_list.item(self.file_list.count()-1).setForeground(QColor("#FF3B30"))
        self.update_preview(); self.refresh_table()

    def process_printing(self):
        paths = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if not paths: return
        temp_out = "print_task.pdf"
        try:
            self.engine.create_layout(paths, self.layout_combo.currentText(), temp_out, self.copy_spin.value())
            if self.mode_combo.currentText() == "打印为PDF":
                save_p, _ = QFileDialog.getSaveFileName(self, "保存PDF", "", "*.pdf")
                if save_p:
                    if os.path.exists(save_p): os.remove(save_p)
                    shutil.move(temp_out, save_p)
                    for i in range(self.file_list.count()): self.engine.save_ledger(self.file_list.item(i).data(Qt.ItemDataRole.UserRole))
                    self.refresh_table(); os.startfile(save_p)
        except Exception as e: QMessageBox.critical(self, "错误", str(e))

    def on_layout_changed(self): self.save_user_config(); self.update_preview()
    def on_copy_changed(self): self.save_user_config(); self.update_preview()
    def add_files(self): 
        f, _ = QFileDialog.getOpenFileNames(self, "选择发票", "", "PDF/OFD (*.pdf *.ofd)")
        if f: self.handle_files(f)
    def add_folder(self): 
        d = QFileDialog.getExistingDirectory(self)
        if d: self.handle_files([os.path.join(d, x) for x in os.listdir(d) if x.lower().endswith(('.pdf', '.ofd'))])
    def remove_selected(self):
        for i in self.file_list.selectedItems(): self.file_list.takeItem(self.file_list.row(i))
        self.update_preview(); self.refresh_table()
    def clear_all(self): self.file_list.clear(); self.update_preview()
    def dragEnterEvent(self, e): e.accept() if e.mimeData().hasUrls() else e.ignore()
    def dropEvent(self, e): self.handle_files([u.toLocalFile() for u in e.mimeData().urls()])
    def load_user_config(self):
        self.layout_combo.setCurrentText(self.settings.value("page_layout", "1×2"))
        self.copy_spin.setValue(self.settings.value("print_copies", 2, type=int))
    def save_user_config(self):
        self.settings.setValue("page_layout", self.layout_combo.currentText())
        self.settings.setValue("print_copies", self.copy_spin.value())

if __name__ == "__main__":
    app = QApplication(sys.argv); app.setStyle("Fusion"); win = ZZPrinterApp(); win.show(); sys.exit(app.exec())