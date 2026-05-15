import os
import sys
import tempfile
import json
import fitz  # PyMuPDF
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QLabel, QPushButton, QListWidget, 
                               QTabWidget, QScrollArea, QFrame, QTextEdit, 
                               QCheckBox, QFileDialog, QMessageBox, QSplitter, 
                               QInputDialog, QLineEdit, QGroupBox, QMenu)
from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QPixmap, QImage, QIcon, QAction, QDesktopServices

# ================= 跨平台处理核心逻辑 =================
IS_WINDOWS = sys.platform.startswith('win')

if IS_WINDOWS:
    try:
        import win32com.client
    except ImportError:
        print("警告: 缺少 pywin32 库。")

def get_system_font():
    font_paths = []
    if IS_WINDOWS:
        font_paths = [
            "C:\\Windows\\Fonts\\simsun.ttc",
            "C:\\Windows\\Fonts\\msyh.ttc",
            "C:\\Windows\\Fonts\\simhei.ttf"
        ]
    elif sys.platform == 'darwin':
        font_paths = [
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
            "/Library/Fonts/Arial Unicode.ttf"
        ]
    for path in font_paths:
        if os.path.exists(path): return path
    return None

SYS_FONT_PATH = get_system_font()
# ======================================================

DEFAULT_CATALOGS = {
    "民商事": [
        "1. *收案登记审批表", "2. *收费凭证", "3. *委托材料（委托合同、风险告知书、授权委托书、公函复印件）", 
        "4. *起诉状、上诉状或答辩状", "5. 阅卷笔录", "6. 委托人谈话笔录", "7. 证据材料", 
        "8. 诉讼（证据、先行）保全申请书、法院裁定书", "9. 承办律师代理意见", "10. 集体讨论记录", 
        "11. 代理词、辩护词", "12. 出庭通知书（传票）", "13. 庭审笔录", 
        "14. *法院通知书、判决书、裁定书、调解书、上诉书", "15. 其他（判决书、裁定书送达证明文件）", 
        "16. *办案质量监督卡", "17. *结案登记表"
    ],
    "刑事": [
        "1. *收案登记审批表", "2. *收费凭证", "3. *委托材料（委托合同、风险告知书、授权委托书、公函复印件或人民法院指定书）", 
        "4. *阅卷笔录", "5. 律师会见被告人、委托人、证人笔录", "6. 律师对此案的调查材料", 
        "7. 律师辩护意见或代理意见", "8. 律师事务所对律师代理的集体讨论意见", "9. *案件起诉书或上诉书", 
        "10. 律师辩护词或代理词", "11. 出庭通知书（传票）", "12. *法院判决书、裁定书", 
        "13. 当事人上诉书和人民检察院抗诉书（如有）", "14. 律师办案小结", "15. 其他（判决书、裁定书送达证明文件）", 
        "16. *办案质量监督卡", "17. *结案登记表"
    ]
}

SUPPORTED_EXTENSIONS = (
    "所有支持的文件 (*.pdf *.doc *.docx *.xls *.xlsx *.ppt *.pptx *.txt *.rtf *.jpg *.jpeg *.png *.bmp *.tif *.tiff);;"
    "PDF 文件 (*.pdf);;"
    "Office/文本 文件 (*.doc *.docx *.xls *.xlsx *.ppt *.pptx *.txt *.rtf);;"
    "图片文件 (*.jpg *.jpeg *.png *.bmp *.tif *.tiff)"
)

VALID_EXTENSIONS_LIST = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.rtf', '.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff']


# ================= 新增：支持拖拽上传的自定义 QListWidget =================
class DropListWidget(QListWidget):
    def __init__(self, unique_key, app_ref, category_name):
        super().__init__()
        self.setAcceptDrops(True) # 开启拖拽接受
        self.unique_key = unique_key
        self.app_ref = app_ref
        self.category_name = category_name
        self.setStyleSheet("QListWidget { background-color: #ffffff; border: 1px solid #dcdcdc; border-radius: 4px; padding: 2px; }"
                           "QListWidget::item:selected { background-color: #e0f7fa; color: #006064; }")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            files = [url.toLocalFile() for url in urls if url.isLocalFile()]
            
            # 过滤支持的文件
            valid_files = [f for f in files if any(f.lower().endswith(ext) for ext in VALID_EXTENSIONS_LIST)]
            
            if valid_files:
                for f in valid_files:
                    self.app_ref.files_data[self.unique_key].append(f)
                    self.addItem(os.path.basename(f))
                self.app_ref.update_stats()
                self.app_ref.log(f"📥 拖拽上传：为 [{self.category_name}] 快速添加了 {len(valid_files)} 个文件。")
            else:
                self.app_ref.log("⚠️ 拖拽失败：包含不支持的文件格式。")
            
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


class ArchiveApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("律师案件归档神器 V7.1 (拖拽修复版)")
        self.resize(1200, 800)
        self.setMinimumSize(1000, 600)
        
        self.files_data = {} 
        self.listbox_dict = {}
        self.reverse_listbox_map = {} 
        
        if SYS_FONT_PATH:
            pdfmetrics.registerFont(TTFont('SystemFont', SYS_FONT_PATH))
            self.report_font = 'SystemFont'
        else:
            pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
            self.report_font = 'STSong-Light'
        
        self.create_menu()
        self.create_ui()

    def create_menu(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu("📂 归档工程管理")
        
        save_action = QAction("💾 保存当前进度 (草稿)", self)
        save_action.triggered.connect(self.save_draft)
        file_menu.addAction(save_action)
        
        load_action = QAction("📂 加载历史进度", self)
        load_action.triggered.connect(self.load_draft)
        file_menu.addAction(load_action)
        
        file_menu.addSeparator()
        exit_action = QAction("❌ 退出程序", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        help_menu = menubar.addMenu("❓ 帮助")
        about_action = QAction("ℹ️ 关于软件", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

    def show_about_dialog(self):
        about_text = """
        <h3>律师案件归档神器</h3>
        <p><b>软件版本：</b> V7.1 (拖拽修复版)</p>
        <p><b>核心更新：</b></p>
        <ul>
            <li>修复多目录预览选中失效的 BUG</li>
            <li>支持原生文件<b>拖拽上传</b></li>
            <li>支持<b>双击</b>列表中的文件调用系统软件直接打开</li>
        </ul>
        
        """
        QMessageBox.about(self, "关于软件", about_text)

    def create_ui(self):
        main_splitter = QSplitter(Qt.Horizontal)
        self.setCentralWidget(main_splitter)

        # ====== 左侧：操作区 ======
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(10, 10, 5, 10)

        self.tab_widget = QTabWidget()
        left_layout.addWidget(self.tab_widget, stretch=1)

        self.tab_civil = QWidget()
        self.tab_criminal = QWidget()
        self.tab_custom = QWidget()

        self.tab_widget.addTab(self.tab_civil, "⚖️ 民商案件")
        self.tab_widget.addTab(self.tab_criminal, "🛡️ 刑事案件")
        self.tab_widget.addTab(self.tab_custom, "📝 自定义目录")

        self.build_catalog_ui(self.tab_civil, DEFAULT_CATALOGS["民商事"], "civil")
        self.build_catalog_ui(self.tab_criminal, DEFAULT_CATALOGS["刑事"], "criminal")
        self.build_custom_ui(self.tab_custom)

        bottom_layout = QVBoxLayout()
        self.chk_compress = QCheckBox("🗜️ 开启电子卷宗瘦身 (极限压缩，适合法院网上立案)")
        self.chk_compress.setStyleSheet("color: #d32f2f; font-weight: bold; font-size: 13px;")
        bottom_layout.addWidget(self.chk_compress)

        self.btn_generate = QPushButton("🚀 一键转换并合并归档")
        self.btn_generate.setMinimumHeight(50)
        self.btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; font-size: 16px; border-radius: 5px;")
        self.btn_generate.clicked.connect(self.start_processing)
        bottom_layout.addWidget(self.btn_generate)
        
        left_layout.addLayout(bottom_layout)

        # ====== 右侧：多功能仪表盘 ======
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(5, 10, 10, 10)

        lbl_preview_title = QLabel("👁️ 文件实时预览 (双击左侧列表文件可打开)")
        lbl_preview_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        right_layout.addWidget(lbl_preview_title)

        self.preview_lbl = QLabel("点击左侧文件查看缩略图\n(支持直接将文件拖拽入左侧对应框内)")
        self.preview_lbl.setAlignment(Qt.AlignCenter)
        self.preview_lbl.setStyleSheet("background-color: #f0f0f0; border: 1px dashed #aaaaaa; border-radius: 8px; color: gray;")
        self.preview_lbl.setMinimumHeight(300)
        right_layout.addWidget(self.preview_lbl, stretch=2)

        lbl_stats_title = QLabel("📊 归档状态")
        lbl_stats_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        right_layout.addWidget(lbl_stats_title)

        self.lbl_stats = QLabel("📂 当前未上传任何文件。")
        self.lbl_stats.setStyleSheet("color: #1976D2; font-size: 13px;")
        right_layout.addWidget(self.lbl_stats)

        lbl_log_title = QLabel("📝 处理日志")
        lbl_log_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        right_layout.addWidget(lbl_log_title)

        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet("background-color: #f9f9f9; font-family: Consolas;")
        right_layout.addWidget(self.txt_log, stretch=1)

        main_splitter.addWidget(left_widget)
        main_splitter.addWidget(right_widget)
        main_splitter.setStretchFactor(0, 6)
        main_splitter.setStretchFactor(1, 4)

    def log(self, message):
        self.txt_log.append(message)
        QApplication.processEvents()

    def update_stats(self):
        total = sum(len(files) for files in self.files_data.values())
        self.lbl_stats.setText(f"📂 当前总计已准备文件：{total} 份")

    def save_draft(self):
        if not any(self.files_data.values()):
            QMessageBox.information(self, "提示", "当前没有文件可保存！")
            return
        path, _ = QFileDialog.getSaveFileName(self, "保存归档工程", "", "JSON 草稿 (*.json)")
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    json.dump(self.files_data, f, ensure_ascii=False, indent=4)
                self.log(f"✅ 草稿已保存至: {os.path.basename(path)}")
                QMessageBox.information(self, "成功", "进度已保存！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存失败: {str(e)}")

    def load_draft(self):
        path, _ = QFileDialog.getOpenFileName(self, "加载归档工程", "", "JSON 草稿 (*.json)")
        if path:
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    loaded_data = json.load(f)
                
                self.files_data = loaded_data
                for key, list_widget in self.listbox_dict.items():
                    list_widget.clear()
                    if key in self.files_data:
                        for file_path in self.files_data[key]:
                            list_widget.addItem(os.path.basename(file_path))
                
                self.update_stats()
                self.log(f"🔄 成功加载历史进度: {os.path.basename(path)}")
                QMessageBox.information(self, "成功", "进度已恢复，请检查文件列表！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"加载失败: {str(e)}")

    def build_catalog_ui(self, parent_widget, catalog_list, tab_prefix):
        layout = QVBoxLayout(parent_widget)
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(0,0,0,0)

        for item in catalog_list:
            self._create_catalog_row(container_layout, item, tab_prefix)
            
        container_layout.addStretch() 
        scroll_area.setWidget(container)
        layout.addWidget(scroll_area)

    def _create_catalog_row(self, layout, item_name, tab_prefix):
        unique_key = f"{tab_prefix}_{item_name}" 
        if unique_key not in self.files_data:
            self.files_data[unique_key] = []
        
        group_box = QGroupBox()
        group_layout = QHBoxLayout(group_box)
        group_layout.setContentsMargins(10, 5, 10, 5)

        lbl_title = QLabel(item_name)
        lbl_title.setWordWrap(True)
        lbl_title.setMinimumWidth(180)
        lbl_title.setMaximumWidth(220)
        lbl_title.setStyleSheet("font-weight: bold;")
        group_layout.addWidget(lbl_title)

        # 使用支持拖拽的新版 ListWidget
        list_widget = DropListWidget(unique_key, self, item_name)
        list_widget.setFixedHeight(70)
        self.listbox_dict[unique_key] = list_widget
        self.reverse_listbox_map[id(list_widget)] = unique_key
        
        # ====== 修复核心 BUG：改用 itemClicked 强制刷新预览 ======
        list_widget.itemClicked.connect(lambda item, lw=list_widget: self.on_listbox_select(lw))
        # 增加双击打开系统文件功能
        list_widget.itemDoubleClicked.connect(lambda item, lw=list_widget: self.open_original_file(lw))
        
        group_layout.addWidget(list_widget, stretch=1)

        btn_layout = QVBoxLayout()
        btn_add = QPushButton("添加文件")
        btn_add.clicked.connect(lambda _, uk=unique_key, name=item_name: self.add_files(uk, name))
        btn_layout.addWidget(btn_add)

        sub_btn_layout = QHBoxLayout()
        btn_up = QPushButton("▲")
        btn_up.setToolTip("上移")
        btn_up.clicked.connect(lambda _, uk=unique_key: self.move_item(uk, -1))
        btn_down = QPushButton("▼")
        btn_down.setToolTip("下移")
        btn_down.clicked.connect(lambda _, uk=unique_key: self.move_item(uk, 1))
        sub_btn_layout.addWidget(btn_up)
        sub_btn_layout.addWidget(btn_down)
        btn_layout.addLayout(sub_btn_layout)

        btn_del = QPushButton("删除")
        btn_del.setStyleSheet("color: red;")
        btn_del.clicked.connect(lambda _, uk=unique_key: self.delete_item(uk))
        btn_layout.addWidget(btn_del)

        group_layout.addLayout(btn_layout)
        layout.addWidget(group_box)

    def build_custom_ui(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        
        layout.addWidget(QLabel("请在下方输入目录项（每行代表一个目录项）："))
        self.text_custom = QTextEdit()
        self.text_custom.setText("1. 顾问合同\n2. 法律意见书\n3. 会议纪要")
        self.text_custom.setMaximumHeight(100)
        layout.addWidget(self.text_custom)
        
        btn_generate_custom = QPushButton("生成下方上传目录")
        btn_generate_custom.clicked.connect(self.generate_custom_catalog)
        layout.addWidget(btn_generate_custom)
        
        self.custom_list_container = QWidget()
        self.custom_list_layout = QVBoxLayout(self.custom_list_container)
        self.custom_list_layout.setContentsMargins(0,0,0,0)
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setWidget(self.custom_list_container)
        layout.addWidget(scroll_area, stretch=1)

    def generate_custom_catalog(self):
        while self.custom_list_layout.count():
            item = self.custom_list_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
                
        content = self.text_custom.toPlainText().strip()
        if not content: return
        items = content.split('\n')
        
        for item in items:
            if item.strip():
                self._create_catalog_row(self.custom_list_layout, item.strip(), "custom")
        self.custom_list_layout.addStretch()
        self.log("✅ 自定义目录已生成。")

    def add_files(self, unique_key, category_name):
        files, _ = QFileDialog.getOpenFileNames(self, f"为【{category_name}】选择文件", "", SUPPORTED_EXTENSIONS)
        if files:
            for f in files:
                self.files_data[unique_key].append(f)
                self.listbox_dict[unique_key].addItem(os.path.basename(f))
            self.update_stats()
            self.log(f"📥 按钮上传：为 [{category_name}] 添加了 {len(files)} 个文件。")

    def move_item(self, unique_key, direction):
        list_widget = self.listbox_dict[unique_key]
        current_row = list_widget.currentRow()
        if current_row < 0: return
        new_row = current_row + direction
        
        if 0 <= new_row < list_widget.count():
            self.files_data[unique_key][current_row], self.files_data[unique_key][new_row] = \
                self.files_data[unique_key][new_row], self.files_data[unique_key][current_row]
            item = list_widget.takeItem(current_row)
            list_widget.insertItem(new_row, item)
            list_widget.setCurrentRow(new_row)
            self.on_listbox_select(list_widget)

    def delete_item(self, unique_key):
        list_widget = self.listbox_dict[unique_key]
        current_row = list_widget.currentRow()
        if current_row < 0: return
        list_widget.takeItem(current_row)
        self.files_data[unique_key].pop(current_row)
        self.update_stats()
        self.preview_lbl.setPixmap(QPixmap())
        self.preview_lbl.setText("文件已移除")

    def open_original_file(self, active_list_widget):
        """双击用系统默认程序打开文件"""
        current_row = active_list_widget.currentRow()
        if current_row < 0: return
        lw_id = id(active_list_widget)
        if lw_id not in self.reverse_listbox_map: return
        unique_key = self.reverse_listbox_map[lw_id]
        
        try:
            file_path = self.files_data[unique_key][current_row]
            QDesktopServices.openUrl(QUrl.fromLocalFile(file_path))
        except Exception:
            pass

    def on_listbox_select(self, active_list_widget):
        """修复后的预览逻辑：排他性高亮 + 强制刷新"""
        # 1. 强制清除其他所有列表的选中高亮状态，保证视觉焦点唯一
        for lw in self.listbox_dict.values():
            if lw != active_list_widget:
                lw.clearSelection()

        current_row = active_list_widget.currentRow()
        if current_row < 0: return
        lw_id = id(active_list_widget)
        if lw_id not in self.reverse_listbox_map: return
        unique_key = self.reverse_listbox_map[lw_id]
        
        try:
            file_path = self.files_data[unique_key][current_row]
            self.render_preview(file_path)
        except IndexError:
            pass 

    def render_preview(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        try:
            if ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff']:
                pixmap = QPixmap(file_path)
                self.preview_lbl.setPixmap(pixmap.scaled(self.preview_lbl.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))
            elif ext == '.pdf':
                doc = fitz.open(file_path)
                page = doc[0]
                pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
                img_data = pix.tobytes("ppm")
                pixmap = QPixmap()
                pixmap.loadFromData(img_data)
                self.preview_lbl.setPixmap(pixmap.scaled(self.preview_lbl.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))
                doc.close()
            elif ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.rtf']:
                self.preview_lbl.setText(f"📄\n该文件为 {ext.upper()} 文档\n暂不支持实时预览（合并时自动处理）\n\n💡 提示：双击左侧列表中的文件名\n可直接在您的系统中打开它！")
            else:
                self.preview_lbl.setText("未知文件格式")
        except Exception as e:
            self.preview_lbl.setText(f"预览失败\n(文件可能已被移动或损坏)")

    # ================= 核心处理逻辑 =================
    def start_processing(self):
        current_tab_idx = self.tab_widget.currentIndex()
        if current_tab_idx == 0:
            active_catalog = DEFAULT_CATALOGS["民商事"]
            tab_prefix = "civil"
        elif current_tab_idx == 1:
            active_catalog = DEFAULT_CATALOGS["刑事"]
            tab_prefix = "criminal"
        else:
            content = self.text_custom.toPlainText().strip()
            active_catalog = [item.strip() for item in content.split('\n') if item.strip()]
            tab_prefix = "custom"

        has_files = any(len(self.files_data.get(f"{tab_prefix}_{cat}", [])) > 0 for cat in active_catalog)
        if not has_files:
            QMessageBox.warning(self, "空目录", "请至少在一个目录项下添加文件后再生成！")
            return

        watermark_text, ok = QInputDialog.getText(self, "防伪水印设置", "请输入背景水印文本\n（留空则不加水印）：", QLineEdit.Normal, "")
        if not ok: return 
        
        compress_mode = self.chk_compress.isChecked()

        save_path, _ = QFileDialog.getSaveFileName(self, "选择归档文件保存位置", "新案件归档.pdf", "PDF 文件 (*.pdf)")
        if not save_path: return

        self.btn_generate.setEnabled(False)
        self.btn_generate.setText("处理中，请勿操作...")
        self.txt_log.clear()
        self.log(f"🚀 开始执行合并归档任务... (瘦身压缩模式: {'已开启' if compress_mode else '未开启'})")

        try:
            self.process_and_merge(active_catalog, save_path, tab_prefix, watermark_text.strip(), compress_mode)
            self.log("🎉 任务全部完成！")
            QMessageBox.information(self, "大功告成", f"归档文件已成功生成并保存至：\n{save_path}")
        except Exception as e:
            self.log(f"❌ 发生致命错误: {str(e)}")
            QMessageBox.critical(self, "发生错误", f"处理过程中发生错误：\n{str(e)}")
        finally:
            self.btn_generate.setEnabled(True)
            self.btn_generate.setText("🚀 一键转换并合并归档")

    def process_and_merge(self, catalog_list, save_path, tab_prefix, watermark_text, compress_mode):
        word_app, excel_app, ppt_app = None, None, None
        if IS_WINDOWS:
            try: word_app = win32com.client.DispatchEx("Word.Application"); word_app.Visible = False
            except: self.log("⚠️ 无法唤醒 Word 进程。")
            try: excel_app = win32com.client.DispatchEx("Excel.Application"); excel_app.Visible = False
            except: self.log("⚠️ 无法唤醒 Excel 进程。")
            try: ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
            except: self.log("⚠️ 无法唤醒 PowerPoint 进程。")

        main_pdf = fitz.open()
        toc_data = [] 

        with tempfile.TemporaryDirectory() as temp_dir:
            for i, cat_name in enumerate(catalog_list):
                unique_key = f"{tab_prefix}_{cat_name}"
                files = self.files_data.get(unique_key, [])
                
                if not files: 
                    toc_data.append((cat_name, False, None))
                    continue 

                start_page = main_pdf.page_count + 1
                toc_data.append((cat_name, True, start_page))
                self.log(f"-> 正在处理：{cat_name} (共 {len(files)} 份文件)")

                for file_idx, file_path in enumerate(files):
                    if not os.path.exists(file_path):
                        self.log(f"   ❌ 找不到文件，跳过: {os.path.basename(file_path)}")
                        continue

                    ext = os.path.splitext(file_path)[1].lower()
                    temp_pdf_path = os.path.join(temp_dir, f"temp_{i}_{file_idx}.pdf")

                    if ext == '.pdf':
                        doc = fitz.open(file_path)
                        main_pdf.insert_pdf(doc)
                        doc.close()
                    
                    elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff']:
                        img = Image.open(file_path)
                        if img.mode != 'RGB': img = img.convert('RGB')
                        
                        if compress_mode:
                            img.thumbnail((1200, 1200)) 
                            temp_jpg_path = temp_pdf_path + ".jpg"
                            img.save(temp_jpg_path, "JPEG", quality=60, optimize=True)
                            doc = fitz.open(temp_jpg_path)
                            pdf_bytes = doc.convert_to_pdf()
                            doc.close()
                            pdf_doc = fitz.open("pdf", pdf_bytes)
                            main_pdf.insert_pdf(pdf_doc)
                            pdf_doc.close()
                        else:
                            img.save(temp_pdf_path, "PDF", resolution=100.0)
                            doc = fitz.open(temp_pdf_path)
                            main_pdf.insert_pdf(doc)
                            doc.close()
                    
                    elif ext in ['.doc', '.docx', '.txt', '.rtf']:
                        if IS_WINDOWS and word_app:
                            self.log(f"   ⚙️ 调用 Word 转换: {os.path.basename(file_path)}")
                            doc = word_app.Documents.Open(os.path.abspath(file_path))
                            doc.SaveAs(os.path.abspath(temp_pdf_path), FileFormat=17) 
                            doc.Close()
                            pdf_doc = fitz.open(temp_pdf_path)
                            main_pdf.insert_pdf(pdf_doc)  
                            pdf_doc.close()
                        else:
                            self._create_mock_pdf(temp_pdf_path, os.path.basename(file_path))
                            doc = fitz.open(temp_pdf_path)
                            main_pdf.insert_pdf(doc)
                            doc.close()

                    elif ext in ['.xls', '.xlsx']:
                        if IS_WINDOWS and excel_app:
                            self.log(f"   ⚙️ 调用 Excel 转换: {os.path.basename(file_path)}")
                            wb = excel_app.Workbooks.Open(os.path.abspath(file_path))
                            ws = wb.ActiveSheet
                            ws.PageSetup.Zoom = False
                            ws.PageSetup.FitToPagesWide = 1 
                            ws.PageSetup.FitToPagesTall = False
                            wb.ExportAsFixedFormat(0, os.path.abspath(temp_pdf_path)) 
                            wb.Close(False)
                            pdf_doc = fitz.open(temp_pdf_path)
                            main_pdf.insert_pdf(pdf_doc)
                            pdf_doc.close()
                        else:
                            self._create_mock_pdf(temp_pdf_path, os.path.basename(file_path))
                            doc = fitz.open(temp_pdf_path)
                            main_pdf.insert_pdf(doc)
                            doc.close()

                    elif ext in ['.ppt', '.pptx']:
                        if IS_WINDOWS and ppt_app:
                            self.log(f"   ⚙️ 调用 PowerPoint 转换: {os.path.basename(file_path)}")
                            presentation = ppt_app.Presentations.Open(os.path.abspath(file_path), WithWindow=False)
                            presentation.SaveAs(os.path.abspath(temp_pdf_path), 32) 
                            presentation.Close()
                            pdf_doc = fitz.open(temp_pdf_path)
                            main_pdf.insert_pdf(pdf_doc)
                            pdf_doc.close()
                        else:
                            self._create_mock_pdf(temp_pdf_path, os.path.basename(file_path))
                            doc = fitz.open(temp_pdf_path)
                            main_pdf.insert_pdf(doc)
                            doc.close()
                            
                    QApplication.processEvents()

            if word_app: word_app.Quit()
            if excel_app: excel_app.Quit()
            if ppt_app: ppt_app.Quit()

            self.log("🖨️ 正在打页码水印和生成书签...")
            
            watermark_doc = None
            if watermark_text:
                wm_path = os.path.join(temp_dir, "watermark.pdf")
                c = canvas.Canvas(wm_path, pagesize=A4)
                c.setFont(self.report_font, 60)
                c.setFillColorRGB(0.85, 0.85, 0.85, alpha=0.5) 
                c.translate(A4[0]/2, A4[1]/2)
                c.rotate(45)
                c.drawCentredString(0, 0, watermark_text)
                c.save()
                watermark_doc = fitz.open(wm_path)

            for page_num in range(main_pdf.page_count):
                page = main_pdf[page_num]
                rect = page.rect
                
                if watermark_doc:
                    page.show_pdf_page(rect, watermark_doc, 0)
                
                text = f"第 {page_num + 1} 页"
                p = fitz.Point(rect.width - 60, rect.height - 30)
                if SYS_FONT_PATH:
                    page.insert_text(p, text, fontsize=12, fontfile=SYS_FONT_PATH, color=(0,0,0))
                else:
                    page.insert_text(p, text, fontsize=12, fontname="cjk", color=(0,0,0))

            toc_pdf_path = os.path.join(temp_dir, "toc.pdf")
            self._generate_toc_pdf(toc_data, toc_pdf_path)
            
            toc_doc = fitz.open(toc_pdf_path)
            main_pdf.insert_pdf(toc_doc, start_at=0)
            toc_page_count = toc_doc.page_count 
            toc_doc.close()
            
            outline = [[1, "卷内目录", 1]]
            for name, has_files, start_page in toc_data:
                if has_files and start_page is not None:
                    actual_target_page = start_page + toc_page_count
                    clean_name = name.split(" ", 1)[-1].replace("*", "").strip() if " " in name else name
                    outline.append([1, clean_name, actual_target_page])
            main_pdf.set_toc(outline)
            
            if watermark_doc: watermark_doc.close()

            self.log("💾 正在保存最终文件...")
            if compress_mode:
                main_pdf.save(save_path, deflate=True, garbage=4)
                self.log("✅ 卷宗瘦身完毕！")
            else:
                main_pdf.save(save_path, deflate=True, garbage=1)
                
            main_pdf.close()

    def _create_mock_pdf(self, output_path, filename):
        c = canvas.Canvas(output_path, pagesize=A4)
        c.setFont(self.report_font, 14)
        c.drawString(100, 700, f"[Debian 开发模式] 模拟转换文件：")
        c.drawString(100, 670, filename)
        c.drawString(100, 640, "（由于在非 Windows 环境，多页文档会被压缩展示为这一页）")
        c.save()

    def _generate_toc_pdf(self, toc_data, output_path):
        c = canvas.Canvas(output_path, pagesize=A4)
        width, height = A4
        c.setFont(self.report_font, 20)
        c.drawCentredString(width/2.0, height - 80, "卷 内 目 录")
        c.setFont(self.report_font, 12)
        y = height - 130
        c.drawString(70, y, "序号 / 内容")
        c.drawString(380, y, "有")
        c.drawString(430, y, "无")
        c.drawString(480, y, "起始页码")
        y -= 15
        c.line(60, y, 540, y) 
        y -= 25
        
        for name, has_files, page_num in toc_data:
            display_name = name if len(name) < 22 else name[:21] + "..."
            c.drawString(70, y, display_name)
            box_size = 10
            c.rect(378, y-1, box_size, box_size) 
            c.rect(428, y-1, box_size, box_size) 
            
            if has_files:
                c.saveState()
                c.setLineWidth(1.5)
                c.line(380, y+3, 383, y)     
                c.line(383, y, 389, y+8)     
                c.restoreState()
                c.drawString(485, y, f"{page_num}")
            else:
                c.saveState()
                c.setLineWidth(1.5)
                c.line(430, y+3, 433, y)
                c.line(433, y, 439, y+8)
                c.restoreState()
                c.drawString(495, y, "-") 
            y -= 25
            if y < 80:
                c.showPage()
                c.setFont(self.report_font, 12)
                y = height - 80
        c.save()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ArchiveApp()
    window.show()
    sys.exit(app.exec())
