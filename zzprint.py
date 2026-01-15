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
from PyQt6.QtGui import QImage, QPixmap, QColor, QKeySequence

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
                    # 数据结构兼容性补丁
                    for key in data:
                        if "items" not in data[key]:
                            data[key]["items"] = [{
                                "项目名称": data[key].get("项目名称", "未知"),
                                "规格型号": data[key].get("规格型号", "无"),
                                "单位": data[key].get("单位", "无"),
                                "数量": data[key].get("数量", "0"),
                                "单价": data[key].get("单价", "0"),
                                "金额": data[key].get("金额", "0.00"),
                                "税率": data[key].get("税率", "0%"),
                                "税额": data[key].get("税额", "0.00"),
                                "合计": data[key].get("合计", "0.00")
                            }]
                        if "处理日期" not in data[key]:
                            data[key]["处理日期"] = "未知"
                    return data
            except: return {}
        return {}

    def save_ledger(self, info):
        if info["发票号码"] != "未知":
            # 更新处理日期：同一张票多次打印，只保留最后一次打印的时间
            info["处理日期"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            self.ledger[info["发票号码"]] = info
            with open(self.ledger_file, 'w', encoding='utf-8') as f:
                json.dump(self.ledger, f, ensure_ascii=False, indent=4)

    def parse_invoice(self, file_path):
        base_info = {
            "发票号码": "未知", "开票日期": "未知", "自产农产品销售": "否",
            "购买方名称": "未知", "购买方税号": "未知",
            "销售方名称": "未知", "销售方税号": "未知",
            "备注": "无", "文件名": os.path.basename(file_path),
            "处理日期": "待处理",
            "items": []
        }
        try:
            with pdfplumber.open(file_path) as pdf:
                text = pdf.pages[0].extract_text()
                lines = text.split('\n')
                
                m_no = re.search(r'发票号码[:：]\s*(\d+)', text)
                if m_no: base_info["发票号码"] = m_no.group(1)
                
                m_date = re.search(r'开票日期[:：]\s*(\d{4}年\d{1,2}月\d{1,2}日)', text)
                if m_date: base_info["开票日期"] = m_date.group(1)
                
                if "自产农产品销售" in text: base_info["自产农产品销售"] = "是"

                names = re.findall(r'名称[:：]\s*([^\n\s]+)', text)
                ids = re.findall(r'纳税人识别号[:：]\s*([A-Z0-9]+)', text)
                if len(names) >= 2: base_info["购买方名称"], base_info["销售方名称"] = names[0], names[1]
                if len(ids) >= 2: base_info["购买方税号"], base_info["销售方税号"] = ids[0], ids[1]

                for line in lines:
                    if '*' in line and any(c.isdigit() for c in line):
                        parts = line.split()
                        if len(parts) >= 6:
                            try:
                                amt_str = parts[-3].replace(',', '').replace('￥','').replace('¥','')
                                tax_str = parts[-1].replace(',', '')
                                amt = float(amt_str)
                                tax = 0.00 if '***' in tax_str or '免税' in parts[-2] else float(tax_str)
                                base_info["items"].append({
                                    "项目名称": parts[0],
                                    "规格型号": parts[1] if len(parts) >= 8 else "无",
                                    "单位": parts[2] if len(parts) >= 8 else (parts[1] if len(parts) == 7 else "无"),
                                    "数量": parts[-5], "单价": parts[-4],
                                    "金额": f"{amt:.2f}", "税率": parts[-2],
                                    "税额": f"{tax:.2f}", "合计": f"{(amt + tax):.2f}"
                                })
                            except: pass
                
                if not base_info["items"]:
                    total_m = re.search(r'[（\(]小写[）\)]\s*[￥¥]?\s*([\d\.]+)', text)
                    if total_m:
                        val = total_m.group(1)
                        base_info["items"].append({"项目名称": "总计", "数量": "1", "金额": val, "税额": "0.00", "合计": val})
        except: pass
        return base_info

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

class ZZPrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.engine = PrintingEngine()
        self.settings = QSettings("ZZStudio", "ZZPrinter")
        self.setWindowTitle("ZZ打票大叔捣腾版 - 3.5 by agentenal")
        self.setMinimumSize(1260, 850)
        self.setAcceptDrops(True)
        
        self.group_stat_active = False
        self.summary_level = 1
        self.theme_mode = self.settings.value("theme", "dark")

        self.init_ui()
        self.apply_theme() 
        self.refresh_table()

    def init_ui(self):
        central_widget = QWidget(); self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # --- 左侧控制面板 ---
        side_scroll = QScrollArea(); side_scroll.setFixedWidth(290); side_scroll.setWidgetResizable(True)
        side_content = QFrame(); side_content.setObjectName("SidePanel")
        side_layout = QVBoxLayout(side_content)
        
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("<h2 style='color:#007AFF;'>ZZ 打票</h2>"))
        self.btn_theme = QPushButton(f"🌓 {self.theme_mode.upper()}"); self.btn_theme.setFixedWidth(80)
        self.btn_theme.clicked.connect(self.toggle_theme); header_layout.addWidget(self.btn_theme)
        side_layout.addLayout(header_layout)
        
        for txt, func in [("添加发票文件", self.add_files), ("从文件夹导入", self.add_folder), ("移除选中文件", self.remove_selected)]:
            btn = QPushButton(txt); btn.clicked.connect(func); side_layout.addWidget(btn)
        
        self.btn_remove_dup = QPushButton("一键移除已打印"); self.btn_remove_dup.setEnabled(False)
        self.btn_remove_dup.clicked.connect(self.remove_duplicates); side_layout.addWidget(self.btn_remove_dup)

        btn_excel = QPushButton("导出 Excel 台账"); btn_excel.setStyleSheet("background-color: #34C759; color: white;")
        btn_excel.clicked.connect(self.export_excel); side_layout.addWidget(btn_excel)
        btn_clear = QPushButton("清空队列"); btn_clear.clicked.connect(self.clear_all); side_layout.addWidget(btn_clear)

        side_layout.addSpacing(15)
        self.mode_combo = QComboBox(); self.mode_combo.addItems(["直接打印", "打印为PDF"])
        side_layout.addWidget(QLabel("打印模式:")); side_layout.addWidget(self.mode_combo)
        self.layout_combo = QComboBox(); self.layout_combo.addItems(["1×1", "1×2", "1×3", "2×2", "2×3", "2×4"])
        self.layout_combo.currentTextChanged.connect(self.update_preview)
        side_layout.addWidget(QLabel("页面布局:")); side_layout.addWidget(self.layout_combo)
        self.copy_spin = QSpinBox(); self.copy_spin.setRange(1, 4); self.copy_spin.setValue(1)
        side_layout.addWidget(QLabel("单张打印份数:")); side_layout.addWidget(self.copy_spin)

        side_layout.addStretch()
        self.log_area = QTextEdit(); self.log_area.setFixedHeight(60); self.log_area.setReadOnly(True); side_layout.addWidget(self.log_area)
        self.btn_print = QPushButton("开始处理 / 打印"); self.btn_print.setFixedHeight(45); self.btn_print.clicked.connect(self.process_printing)
        side_layout.addWidget(self.btn_print)
        btn_quit = QPushButton("退出程序"); btn_quit.setObjectName("QuitBtn"); btn_quit.clicked.connect(self.close); side_layout.addWidget(btn_quit)
        side_scroll.setWidget(side_content); main_layout.addWidget(side_scroll)

        # --- 右侧内容区 ---
        content_layout = QVBoxLayout()
        top_split = QHBoxLayout()
        self.file_list = QListWidget(); self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        v1 = QVBoxLayout(); v1.addWidget(QLabel("<b>待打印队列</b>")); v1.addWidget(self.file_list); top_split.addLayout(v1, 1)
        self.preview_label = QLabel("预览区"); self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scroll_area = QScrollArea(); self.scroll_area.setWidget(self.preview_label); self.scroll_area.setWidgetResizable(True)
        v2 = QVBoxLayout(); v2.addWidget(QLabel("<b>实时预览</b>")); v2.addWidget(self.scroll_area); top_split.addLayout(v2, 1)
        content_layout.addLayout(top_split, 2)

        # --- 增强型筛选工具栏 ---
        filter_box = QFrame(); filter_box.setFixedHeight(45)
        filter_layout = QHBoxLayout(filter_box); filter_layout.setContentsMargins(0,0,0,0)
        filter_layout.addWidget(QLabel("<b>销售方:</b>"))
        self.search_seller = QLineEdit(); self.search_seller.setPlaceholderText("关键词...")
        self.search_seller.textChanged.connect(self.refresh_table); filter_layout.addWidget(self.search_seller, 1)
        
        filter_layout.addWidget(QLabel("<b>开票日期:</b>"))
        self.search_date = QLineEdit(); self.search_date.setPlaceholderText("2026-01"); self.search_date.setFixedWidth(80)
        self.search_date.textChanged.connect(self.refresh_table); filter_layout.addWidget(self.search_date)

        filter_layout.addWidget(QLabel("<b>处理日期:</b>"))
        self.search_proc_date = QLineEdit(); self.search_proc_date.setPlaceholderText("天 或 区间(至)"); self.search_proc_date.setFixedWidth(140)
        self.search_proc_date.textChanged.connect(self.refresh_table); filter_layout.addWidget(self.search_proc_date)

        self.btn_group_stat = QPushButton("📊 分组: 关"); self.btn_group_stat.setCheckable(True)
        self.btn_group_stat.clicked.connect(self.toggle_group_stat); filter_layout.addWidget(self.btn_group_stat)
        self.btn_sum_level = QPushButton("汇总: 一级"); self.btn_sum_level.clicked.connect(self.toggle_sum_level)
        filter_layout.addWidget(self.btn_sum_level)
        btn_reset = QPushButton("重置"); btn_reset.clicked.connect(self.reset_filters); filter_layout.addWidget(btn_reset)
        content_layout.addWidget(filter_box)

        self.table = CopyableTable()
        headers = ["发票号码", "开票日期", "处理日期", "销售方", "明细项目", "数量", "税额", "金额", "价税合计"]
        self.table.setColumnCount(len(headers)); self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        content_layout.addWidget(self.table, 1)
        main_layout.addLayout(content_layout)

    def toggle_theme(self):
        self.theme_mode = "light" if self.theme_mode == "dark" else "dark"
        self.settings.setValue("theme", self.theme_mode); self.btn_theme.setText(f"🌓 {self.theme_mode.upper()}"); self.apply_theme()
    def toggle_group_stat(self):
        self.group_stat_active = self.btn_group_stat.isChecked()
        self.btn_group_stat.setText("📊 分组: 开" if self.group_stat_active else "📊 分组: 关"); self.refresh_table()
    def toggle_sum_level(self):
        self.summary_level = 2 if self.summary_level == 1 else 1
        self.btn_sum_level.setText("汇总: 一级" if self.summary_level == 1 else "汇总: 二级")
        if self.group_stat_active: self.refresh_table()

    def get_filtered_raw_list(self):
        """ 获取满足筛选条件的原始发票基础信息列表 """
        seller_key = self.search_seller.text().strip().lower()
        date_key = self.search_date.text().strip().replace("-", "").replace("年","").replace("月","")
        proc_date_key = self.search_proc_date.text().strip()
        
        results = []
        for no, base in self.engine.ledger.items():
            # 开票日期和销售方匹配
            if seller_key in base.get("销售方名称", "").lower() and date_key in base.get("开票日期", "").replace("年","").replace("月",""):
                
                # 处理日期筛选
                p_date = base.get("处理日期", "未知")
                match_proc = True
                if proc_date_key:
                    if "至" in proc_date_key:
                        try:
                            start_s, end_s = proc_date_key.split("至")
                            cur_d = p_date.split(" ")[0]
                            match_proc = (start_s.strip() <= cur_d <= end_s.strip())
                        except: match_proc = False
                    else:
                        match_proc = (proc_date_key in p_date)
                
                if match_proc:
                    results.append(base)
        return results

    def refresh_table(self):
        self.table.setRowCount(0)
        filtered_bases = self.get_filtered_raw_list()
        
        flat_data = []
        for base in filtered_bases:
            for item in base.get("items", []):
                flat_data.append({**base, **item})
        
        if not flat_data: return
        df = pd.DataFrame(flat_data)
        for col in ['合计', '金额', '税额']: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        if not self.group_stat_active: 
            self._fill_table_rows(df.to_dict('records'))
        else:
            if self.summary_level == 1:
                for seller, g_s in df.groupby('销售方名称', sort=False):
                    self._fill_table_rows(g_s.to_dict('records'))
                    self._add_summary_row(f"【{seller}】小计", g_s['税额'].sum(), g_s['金额'].sum(), g_s['合计'].sum(), QColor("#E1F5FE" if self.theme_mode == "light" else "#01579B"))
            else:
                for (seller, inv_no), g_inv in df.groupby(['销售方名称', '发票号码'], sort=False):
                    self._fill_table_rows(g_inv.to_dict('records'))
                    self._add_summary_row(f"票号 {inv_no} 小计", g_inv['税额'].sum(), g_inv['金额'].sum(), g_inv['合计'].sum(), QColor("#F1F8E9" if self.theme_mode == "light" else "#1B5E20"))

    def _fill_table_rows(self, rows):
        for r_data in rows:
            r = self.table.rowCount(); self.table.insertRow(r)
            vals = [r_data["发票号码"], r_data["开票日期"], r_data.get("处理日期","未知"), r_data["销售方名称"], r_data["项目名称"], r_data["数量"], f"{r_data['税额']:.2f}", f"{r_data['金额']:.2f}", f"{r_data['合计']:.2f}"]
            for i, v in enumerate(vals): self.table.setItem(r, i, QTableWidgetItem(str(v)))

    def _add_summary_row(self, label, tax, amt, total, color):
        r = self.table.rowCount(); self.table.insertRow(r)
        sum_item = QTableWidgetItem(label); sum_item.setBackground(color)
        self.table.setItem(r, 3, sum_item)
        for i, val in [(6, tax), (7, amt), (8, total)]:
            ti = QTableWidgetItem(f"{val:.2f}"); ti.setBackground(color); ti.setForeground(QColor("#FF9500"))
            self.table.setItem(r, i, ti)

    def export_excel(self):
        """ 导出台账：彻底展平明细，每个字段独立一列 """
        filtered_bases = self.get_filtered_raw_list()
        if not filtered_bases:
            QMessageBox.warning(self, "提示", "没有数据可导出！")
            return

        # 定义 Excel 列顺序
        col_order = [
            "发票号码", "开票日期", "处理日期", "销售方名称", "销售方税号", "自产农产品销售",
            "项目名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额", "合计",
            "购买方名称", "购买方税号", "备注"
        ]

        p, _ = QFileDialog.getSaveFileName(self, "保存台账", f"发票台账_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx", "*.xlsx")
        if p:
            try:
                # 核心逻辑：循环发票 -> 循环明细 -> 每一行明细生成一个 DataFrame 行
                export_data = []
                for base in filtered_bases:
                    items = base.get("items", [])
                    for item in items:
                        # 合并基础字段和明细字段
                        row = {**base, **item}
                        export_data.append(row)
                
                df = pd.DataFrame(export_data)
                # 补全可能缺失的列
                for c in col_order:
                    if c not in df.columns: df[c] = ""
                
                # 按照指定顺序导出，并排除掉 items 列表原文列
                df[col_order].to_excel(p, index=False)
                self.log_area.append(f"导出成功！共 {len(export_data)} 行明细。")
                # os.startfile(os.path.dirname(p)) 默认导出后不打开文件夹
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")

    def process_printing(self):
        paths = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if not paths: return
        out = "print_task.pdf"
        self.engine.create_layout(paths, self.layout_combo.currentText(), out, self.copy_spin.value())
        save_p, _ = QFileDialog.getSaveFileName(self, "保存打印文件", "", "*.pdf")
        if save_p:
            shutil.move(out, save_p)
            # 处理每一张发票并记录时间
            for i in range(self.file_list.count()): 
                info = self.file_list.item(i).data(Qt.ItemDataRole.UserRole)
                self.engine.save_ledger(info)
            self.refresh_table(); os.startfile(save_p)

    def remove_duplicates(self):
        for i in range(self.file_list.count() - 1, -1, -1):
            if self.file_list.item(i).data(Qt.ItemDataRole.UserRole)["发票号码"] in self.engine.ledger: self.file_list.takeItem(i)
        self.btn_remove_dup.setEnabled(False); self.update_preview()

    def handle_files(self, paths):
        has_dup = False
        for p in paths:
            if p.lower().endswith(('.pdf', '.ofd')):
                info = self.engine.parse_invoice(p); self.file_list.addItem(p)
                self.file_list.item(self.file_list.count()-1).setData(Qt.ItemDataRole.UserRole, info)
                if info["发票号码"] in self.engine.ledger:
                    self.file_list.item(self.file_list.count()-1).setForeground(QColor("#FF3B30")); has_dup = True
        self.btn_remove_dup.setEnabled(has_dup); self.update_preview()

    def apply_theme(self):
        dark = self.theme_mode == "dark"
        cfg = {"bg": "#1C1C1E" if dark else "#F2F2F7", "panel": "#2C2C2E" if dark else "#FFFFFF", "text": "#FFFFFF" if dark else "#1C1C1E", "border": "#3A3A3C" if dark else "#D1D1D6"}
        self.setStyleSheet(f"QMainWindow, QScrollArea {{ background: {cfg['bg']}; }} QFrame#SidePanel {{ background: {cfg['panel']}; border-radius: 12px; margin: 5px; }} QLabel {{ color: {cfg['text']}; }} QPushButton {{ background: #007AFF; color: white; border-radius: 8px; padding: 6px; border:none; }} QTableWidget {{ background: {cfg['panel']}; color: {cfg['text']}; border: 1px solid {cfg['border']}; }} QPushButton#QuitBtn {{ background: #FF3B30; }}")

    def update_preview(self):
        if self.file_list.count() == 0: self.preview_label.setPixmap(QPixmap()); return
        paths = [self.file_list.item(i).text() for i in range(min(self.file_list.count(), 6))]
        self.engine.create_layout(paths, self.layout_combo.currentText(), "pre.pdf", self.copy_spin.value())
        try:
            doc = fitz.open("pre.pdf")
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(1.2, 1.2))
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
    def reset_filters(self): self.search_seller.clear(); self.search_date.clear(); self.search_proc_date.clear(); self.refresh_table()
    def dragEnterEvent(self, e): e.accept() if e.mimeData().hasUrls() else e.ignore()
    def dropEvent(self, e): self.handle_files([u.toLocalFile() for u in e.mimeData().urls()])

if __name__ == "__main__":
    app = QApplication(sys.argv); app.setStyle("Fusion"); win = ZZPrinterApp(); win.show(); sys.exit(app.exec())