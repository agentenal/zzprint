import sys
import os
import hashlib
import json
import shutil
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QComboBox, QListWidget, 
                             QLabel, QTextEdit, QFileDialog, QFrame, QSpinBox, 
                             QMessageBox, QScrollArea, QAbstractItemView)
from PyQt6.QtCore import Qt, QRectF, QSettings
from PyQt6.QtGui import QImage, QPixmap, QPainter, QPalette

# --- 核心处理引擎 ---
class PrintingEngine:
    def __init__(self, history_file="print_history.json"):
        self.history_file = history_file
        self.history = self.load_history()
        self.layout_map = {
            "1×1": (1, 1), "1×2": (2, 1), "1×3": (3, 1),
            "2×2": (2, 2), "2×3": (3, 2), "2×4": (4, 2)
        }

    def load_history(self):
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except: return []
        return []

    def save_history(self, file_hash):
        if file_hash not in self.history:
            self.history.append(file_hash)
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.history, f)

    def get_file_md5(self, file_path):
        hasher = hashlib.md5()
        with open(file_path, 'rb') as f:
            while chunk := f.read(8192):
                hasher.update(chunk)
        return hasher.hexdigest()

    def create_layout(self, input_files, layout_desc, output_path, copies=1):
        a4_w, a4_h = 595, 842 
        doc = fitz.open()
        rows, cols = self.layout_map.get(layout_desc, (1, 1))
        files_per_page = rows * cols
        expanded_files = [f for f in input_files for _ in range(copies)]

        for i in range(0, len(expanded_files), files_per_page):
            new_page = doc.new_page(width=a4_w, height=a4_h)
            batch = expanded_files[i:i + files_per_page]
            cell_w, cell_h = a4_w / cols, a4_h / rows
            for idx, file_path in enumerate(batch):
                try:
                    src_doc = fitz.open(file_path)
                    r, c = divmod(idx, cols)
                    target_rect = fitz.Rect(c*cell_w+10, r*cell_h+10, (c+1)*cell_w-10, (r+1)*cell_h-10)
                    new_page.show_pdf_page(target_rect, src_doc, 0)
                    src_doc.close()
                except: pass
        doc.save(output_path)
        doc.close()

# --- 主界面 ---
class ZZPrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.engine = PrintingEngine()
        self.settings = QSettings("ZZStudio", "ZZPrinter")
        self.setWindowTitle("ZZ打票 - 大叔捣腾的小工具")
        self.setMinimumSize(1050, 800)
        
        # 开启拖拽支持
        self.setAcceptDrops(True)
        
        self.init_ui()
        self.load_user_config() # 加载历史配置
        self.apply_theme()

    def is_dark_mode(self):
        palette = QApplication.palette()
        return palette.color(QPalette.ColorRole.Window).lightness() < 128

    def apply_theme(self):
        dark = self.is_dark_mode()
        cfg = {
            "bg": "#1C1C1E" if dark else "#F2F2F7",
            "panel": "#2C2C2E" if dark else "#FFFFFF",
            "text": "#FFFFFF" if dark else "#1C1C1E",
            "border": "#3A3A3C" if dark else "#D1D1D6",
            "item_bg": "#3A3A3C" if dark else "#FFFFFF",
            "btn_sec": "#3A3A3C" if dark else "#E5E5EA",
            "log_bg": "#1C1C1E" if dark else "#F2F2F7"
        }
        self.setStyleSheet(f"""
            QMainWindow {{ background-color: {cfg['bg']}; }}
            QLabel {{ color: {cfg['text']}; font-family: "Microsoft YaHei"; }} 
            QFrame#SidePanel {{ background-color: {cfg['panel']}; border-radius: 20px; margin: 5px; }}
            QComboBox, QSpinBox {{
                border: 1px solid {cfg['border']}; border-radius: 8px; padding: 5px;
                background: {cfg['item_bg']}; color: {cfg['text']}; min-height: 25px;
            }}
            QPushButton {{
                background-color: #007AFF; color: white; border-radius: 12px;
                padding: 10px; font-weight: bold; font-size: 13px;
            }}
            QPushButton#ActionBtn {{ background-color: {cfg['btn_sec']}; color: #007AFF; }}
            QPushButton#QuitBtn {{ background-color: #FFEBEE; color: #FF3B30; border: 1px solid #FFCDD2; }}
            QListWidget {{ 
                background: {cfg['item_bg']}; border-radius: 12px; border: 1px solid {cfg['border']}; 
                color: {cfg['text']}; font-size: 12px; outline: none;
            }}
            QTextEdit {{ background: {cfg['log_bg']}; border-radius: 10px; color: {cfg['text']}; font-size: 10px; border: none; }}
            QScrollArea {{ border-radius: 15px; background-color: #000000; border: 1px solid {cfg['border']}; }}
        """)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # --- 左侧 ---
        side_panel = QFrame(); side_panel.setObjectName("SidePanel")
        side_panel.setFixedWidth(280)
        side_layout = QVBoxLayout(side_panel)
        side_layout.addWidget(QLabel("<h2 style='color:#007AFF;'>ZZ 打票</h2>"))
        
        btn_add = QPushButton("添加发票文件")
        btn_add.clicked.connect(self.add_files)
        side_layout.addWidget(btn_add)
        
        btn_folder = QPushButton("从文件夹导入")
        btn_folder.setObjectName("ActionBtn")
        btn_folder.clicked.connect(self.add_folder)
        side_layout.addWidget(btn_folder)

        btn_remove = QPushButton("移除选中文件")
        btn_remove.setObjectName("ActionBtn")
        btn_remove.clicked.connect(self.remove_selected)
        side_layout.addWidget(btn_remove)
        
        btn_clear = QPushButton("清空所有列表")
        btn_clear.setObjectName("ActionBtn")
        btn_clear.clicked.connect(self.clear_all)
        side_layout.addWidget(btn_clear)

        side_layout.addSpacing(15)
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["直接打印", "打印为PDF"])
        self.mode_combo.currentTextChanged.connect(self.save_user_config)
        side_layout.addWidget(QLabel("打印模式:"))
        side_layout.addWidget(self.mode_combo)

        self.layout_combo = QComboBox()
        self.layout_combo.addItems(["1×1", "1×2", "1×3", "2×2", "2×3", "2×4"])
        self.layout_combo.currentTextChanged.connect(self.on_layout_changed)
        side_layout.addWidget(QLabel("页面布局:"))
        side_layout.addWidget(self.layout_combo)

        self.copy_spin = QSpinBox(); self.copy_spin.setRange(1, 4)
        self.copy_spin.valueChanged.connect(self.on_copy_changed)
        side_layout.addWidget(QLabel("单张打印份数:"))
        side_layout.addWidget(self.copy_spin)

        side_layout.addStretch()
        self.log_area = QTextEdit(); self.log_area.setFixedHeight(80); self.log_area.setReadOnly(True)
        side_layout.addWidget(self.log_area)

        self.btn_print = QPushButton("开始打印")
        self.btn_print.setFixedHeight(45)
        self.btn_print.clicked.connect(self.process_printing)
        side_layout.addWidget(self.btn_print)

        btn_quit = QPushButton("退出程序")
        btn_quit.setObjectName("QuitBtn")
        btn_quit.setFixedHeight(40)
        btn_quit.clicked.connect(self.close)
        side_layout.addWidget(btn_quit)

        # --- 右侧 ---
        content_layout = QVBoxLayout()
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        content_layout.addWidget(QLabel("<b>发票列表 (支持拖拽/多选)</b>"))
        content_layout.addWidget(self.file_list, 1)

        content_layout.addWidget(QLabel("<b>实时预览</b>"))
        self.scroll_area = QScrollArea(); self.scroll_area.setWidgetResizable(True)
        self.preview_label = QLabel("将文件拖拽至此处添加"); self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scroll_area.setWidget(self.preview_label)
        content_layout.addWidget(self.scroll_area, 2)

        main_layout.addWidget(side_panel)
        main_layout.addLayout(content_layout)

    # --- 拖拽实现 ---
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        added_count = 0
        for f in files:
            if f.lower().endswith(('.pdf', '.ofd')):
                self.file_list.addItem(f)
                added_count += 1
        if added_count > 0:
            self.log_area.append(f"通过拖拽添加了 {added_count} 个文件")
            self.update_preview()

    # --- 配置保存/加载 ---
    def load_user_config(self):
        """恢复上次运行的设置，如果没有则使用默认值"""
        mode = self.settings.value("print_mode", "打印为PDF")
        layout = self.settings.value("page_layout", "1×2")
        copies = self.settings.value("print_copies", 2, type=int)
        
        self.mode_combo.setCurrentText(mode)
        self.layout_combo.setCurrentText(layout)
        self.copy_spin.setValue(copies)

    def save_user_config(self):
        self.settings.setValue("print_mode", self.mode_combo.currentText())
        self.settings.setValue("page_layout", self.layout_combo.currentText())
        self.settings.setValue("print_copies", self.copy_spin.value())

    def on_layout_changed(self):
        self.save_user_config()
        self.update_preview()

    def on_copy_changed(self):
        self.save_user_config()
        self.update_preview()

    # --- 基础逻辑 ---
    def add_files(self):
        last_path = self.settings.value("last_path", os.path.expanduser("~"))
        files, _ = QFileDialog.getOpenFileNames(self, "选择发票", last_path, "PDF/OFD (*.pdf *.ofd)")
        if files:
            self.settings.setValue("last_path", os.path.dirname(files[0]))
            for f in files: self.file_list.addItem(f)
            self.update_preview()

    def add_folder(self):
        last_path = self.settings.value("last_path", os.path.expanduser("~"))
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", last_path)
        if folder:
            self.settings.setValue("last_path", folder)
            for f in os.listdir(folder):
                if f.lower().endswith(('.pdf', '.ofd')): self.file_list.addItem(os.path.join(folder, f))
            self.update_preview()

    def remove_selected(self):
        for item in self.file_list.selectedItems(): self.file_list.takeItem(self.file_list.row(item))
        self.update_preview()

    def clear_all(self):
        self.file_list.clear(); self.update_preview()

    def update_preview(self):
        paths = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if not paths:
            self.preview_label.setPixmap(QPixmap()); self.preview_label.setText("将文件拖拽至此处添加")
            return
        temp_pdf = "preview_cache.pdf"
        try:
            self.engine.create_layout(paths[:12], self.layout_combo.currentText(), temp_pdf, self.copy_spin.value())
            doc = fitz.open(temp_pdf)
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
            self.preview_label.setPixmap(QPixmap.fromImage(img).scaledToWidth(self.scroll_area.width()-40, Qt.TransformationMode.SmoothTransformation))
            doc.close()
        except: pass

    def process_printing(self):
        paths = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if not paths: return
        temp_output = "print_task.pdf"
        try:
            self.engine.create_layout(paths, self.layout_combo.currentText(), temp_output, self.copy_spin.value())
            if self.mode_combo.currentText() == "打印为PDF":
                last_save = self.settings.value("last_save_path", os.path.expanduser("~"))
                save_path, _ = QFileDialog.getSaveFileName(self, "保存PDF", last_save, "*.pdf")
                if save_path:
                    self.settings.setValue("last_save_path", os.path.dirname(save_path))
                    if os.path.exists(save_path): os.remove(save_path)
                    shutil.move(temp_output, save_path)
                    os.startfile(save_path)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ZZPrinterApp()
    window.show()
    sys.exit(app.exec())