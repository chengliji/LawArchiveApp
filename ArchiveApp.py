import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import os
import sys
import tempfile
import json
import fitz  # PyMuPDF
from PIL import Image, ImageTk
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4

# ================= 跨平台处理核心逻辑 =================
IS_WINDOWS = sys.platform.startswith('win')

if IS_WINDOWS:
    try:
        import win32com.client
    except ImportError:
        print("警告: 缺少 pywin32 库。")

def get_windows_font():
    if not IS_WINDOWS: return None
    font_paths = [
        "C:\\Windows\\Fonts\\simsun.ttc",
        "C:\\Windows\\Fonts\\msyh.ttc",
        "C:\\Windows\\Fonts\\simhei.ttf"
    ]
    for path in font_paths:
        if os.path.exists(path): return path
    return None

SYS_FONT_PATH = get_windows_font()
# ======================================================

DEFAULT_CATALOGS = {
    "民商事": [
        "1. 收案登记审批表", "2. 收费凭证", "3. 委托材料", "4. 起诉状、上诉状或答辩状", 
        "5. 阅卷笔录", "6. 委托人谈话笔录", "7. 证据材料", "8. 诉讼(证据、先行)保全申请书、法院裁定书", 
        "9. 承办律师代理意见", "10. 集体讨论记录", "11. 代理词、辩护词", "12. 出庭通知书(传票)", 
        "13. 庭审笔录", "14. 法院通知书、判决书、裁定书等", "15. 其他送达证明文件", 
        "16. 办案质量监督卡", "17. 结案登记表"
    ],
    "刑事": [
        "1. 收案登记审批表", "2. 收费凭证", "3. 委托材料", "4. 阅卷笔录", 
        "5. 会见被告人、委托人、证人笔录", "6. 调查材料", "7. 辩护或代理意见", "8. 集体讨论意见", 
        "9. 起诉书或上诉书", "10. 辩护词或代理词", "11. 出庭通知书(传票)", "12. 法院判决书、裁定书", 
        "13. 上诉书和抗诉书", "14. 办案小结", "15. 其他送达证明文件", "16. 办案质量监督卡", "17. 结案登记表"
    ]
}

# --- 扩展支持的文件格式 ---
SUPPORTED_EXTENSIONS = (
    "*.pdf *.doc *.docx *.xls *.xlsx *.ppt *.pptx *.txt *.rtf "
    "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"
)

class ArchiveApp:
    def __init__(self, root):
        self.root = root
        self.root.title("律师案件归档神器 V5.0 (全格式兼容版)")
        self.root.geometry("1100x750") 
        
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
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="保存当前进度 (草稿)", command=self.save_draft)
        file_menu.add_command(label="加载历史进度", command=self.load_draft)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        menubar.add_cascade(label="📂 归档工程管理", menu=file_menu)
        self.root.config(menu=menubar)

    def create_ui(self):
        paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=3) 

        self.notebook = ttk.Notebook(left_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tab_civil = ttk.Frame(self.notebook)
        self.tab_criminal = ttk.Frame(self.notebook)
        self.tab_non_litigation = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_civil, text="民商案件")
        self.notebook.add(self.tab_criminal, text="刑事案件")
        self.notebook.add(self.tab_non_litigation, text="非诉案件(自定义)")

        self.build_catalog_ui(self.tab_civil, DEFAULT_CATALOGS["民商事"], "civil")
        self.build_catalog_ui(self.tab_criminal, DEFAULT_CATALOGS["刑事"], "criminal")
        self.build_custom_ui(self.tab_non_litigation)

        bottom_frame = tk.Frame(left_frame)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)
        
        self.compress_var = tk.BooleanVar(value=False)
        chk_compress = tk.Checkbutton(bottom_frame, text="🗜️ 开启电子卷宗瘦身 (极限压缩，适合法院网上立案)", 
                                      variable=self.compress_var, font=("微软雅黑", 10), fg="#d32f2f")
        chk_compress.pack(pady=(0, 5))

        self.btn_generate = tk.Button(bottom_frame, text="🚀 一键转换并合并归档", 
                                      bg="#4CAF50", fg="white", font=("微软雅黑", 12, "bold"), 
                                      command=self.start_processing)
        self.btn_generate.pack(pady=5, fill=tk.X, ipadx=20, ipady=5)

        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=2)

        preview_label = tk.Label(right_frame, text="👁️ 文件实时预览", font=("微软雅黑", 10, "bold"))
        preview_label.pack(anchor="w", pady=(10,0), padx=5)
        
        self.preview_canvas = tk.Canvas(right_frame, bg="#e0e0e0", height=300, relief="sunken", bd=2)
        self.preview_canvas.pack(fill=tk.X, padx=5, pady=5)
        self.preview_canvas.create_text(200, 150, text="点击左侧文件查看缩略图\n(支持 PDF 与 各类图片)", fill="gray", font=("微软雅黑", 10), justify="center", tags="info_text")
        self.current_preview_img = None 

        stats_label = tk.Label(right_frame, text="📊 归档状态统计", font=("微软雅黑", 10, "bold"))
        stats_label.pack(anchor="w", pady=(15,0), padx=5)
        self.lbl_stats = tk.Label(right_frame, text="当前未上传任何文件。", fg="blue", justify="left")
        self.lbl_stats.pack(anchor="w", padx=5, pady=5)

        log_label = tk.Label(right_frame, text="📝 处理日志", font=("微软雅黑", 10, "bold"))
        log_label.pack(anchor="w", pady=(10,0), padx=5)
        self.txt_log = scrolledtext.ScrolledText(right_frame, height=12, state=tk.DISABLED, bg="#f5f5f5")
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def log(self, message):
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.insert(tk.END, message + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state=tk.DISABLED)
        self.root.update()

    def update_stats(self):
        total = sum(len(files) for files in self.files_data.values())
        self.lbl_stats.config(text=f"📂 当前总计已准备文件：{total} 份")

    def save_draft(self):
        if not any(self.files_data.values()):
            messagebox.showinfo("提示", "当前没有文件可保存！")
            return
        path = filedialog.asksaveasfilename(title="保存归档工程", defaultextension=".json", filetypes=[("JSON 草稿", "*.json")])
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    json.dump(self.files_data, f, ensure_ascii=False, indent=4)
                self.log(f"✅ 草稿已保存至: {os.path.basename(path)}")
                messagebox.showinfo("成功", "进度已保存！")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {str(e)}")

    def load_draft(self):
        path = filedialog.askopenfilename(title="加载归档工程", filetypes=[("JSON 草稿", "*.json")])
        if path:
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    loaded_data = json.load(f)
                
                self.files_data = loaded_data
                for key, listbox in self.listbox_dict.items():
                    listbox.delete(0, tk.END)
                    if key in self.files_data:
                        for file_path in self.files_data[key]:
                            listbox.insert(tk.END, os.path.basename(file_path))
                
                self.update_stats()
                self.log(f"🔄 成功加载历史进度: {os.path.basename(path)}")
                messagebox.showinfo("成功", "进度已恢复，请检查文件列表！")
            except Exception as e:
                messagebox.showerror("错误", f"加载失败: {str(e)}")

    def build_catalog_ui(self, parent_frame, catalog_list, tab_prefix):
        canvas_widget = tk.Canvas(parent_frame)
        scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=canvas_widget.yview)
        scrollable_frame = ttk.Frame(canvas_widget)

        scrollable_frame.bind("<Configure>", lambda e: canvas_widget.configure(scrollregion=canvas_widget.bbox("all")))
        canvas_widget.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas_widget.configure(yscrollcommand=scrollbar.set)

        canvas_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _bind_mousewheel(e):
            canvas_widget.bind_all("<MouseWheel>", lambda event: canvas_widget.yview_scroll(int(-1*(event.delta/120)), "units"))
            canvas_widget.bind_all("<Button-4>", lambda event: canvas_widget.yview_scroll(-1, "units"))
            canvas_widget.bind_all("<Button-5>", lambda event: canvas_widget.yview_scroll(1, "units"))

        def _unbind_mousewheel(e):
            canvas_widget.unbind_all("<MouseWheel>")
            canvas_widget.unbind_all("<Button-4>")
            canvas_widget.unbind_all("<Button-5>")

        canvas_widget.bind("<Enter>", _bind_mousewheel)
        canvas_widget.bind("<Leave>", _unbind_mousewheel)

        for item in catalog_list:
            self._create_catalog_row(scrollable_frame, item, tab_prefix)

    def _create_catalog_row(self, parent, item_name, tab_prefix):
        unique_key = f"{tab_prefix}_{item_name}" 
        if unique_key not in self.files_data:
            self.files_data[unique_key] = []
        
        row_frame = tk.Frame(parent, pady=5, padx=10)
        row_frame.pack(fill=tk.X, expand=True)

        tk.Label(row_frame, text=item_name, width=25, anchor="w", font=("微软雅黑", 10)).pack(side=tk.LEFT)
        listbox = tk.Listbox(row_frame, height=3, width=50, exportselection=False) 
        listbox.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        self.listbox_dict[unique_key] = listbox
        self.reverse_listbox_map[str(listbox)] = unique_key
        listbox.bind('<<ListboxSelect>>', self.on_listbox_select)

        btn_frame = tk.Frame(row_frame)
        btn_frame.pack(side=tk.RIGHT)
        tk.Button(btn_frame, text="添加文件", command=lambda: self.add_files(unique_key, item_name)).grid(row=0, column=0, padx=2, pady=2)
        tk.Button(btn_frame, text="上移", command=lambda: self.move_item(unique_key, -1)).grid(row=0, column=1, padx=2, pady=2)
        tk.Button(btn_frame, text="下移", command=lambda: self.move_item(unique_key, 1)).grid(row=1, column=0, padx=2, pady=2)
        tk.Button(btn_frame, text="删除", command=lambda: self.delete_item(unique_key)).grid(row=1, column=1, padx=2, pady=2)

    def build_custom_ui(self, parent_frame):
        top_frame = tk.Frame(parent_frame)
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(top_frame, text="请在下方输入目录项（每行代表一个目录项）：", font=("微软雅黑", 10)).pack(anchor="w")
        self.text_custom = scrolledtext.ScrolledText(top_frame, height=6, width=60)
        self.text_custom.pack(pady=5, fill=tk.X)
        self.text_custom.insert(tk.END, "1. 顾问合同\n2. 法律意见书\n3. 会议纪要")
        tk.Button(top_frame, text="生成下方上传目录", command=self.generate_custom_catalog).pack(pady=5)
        self.custom_list_frame = tk.Frame(parent_frame)
        self.custom_list_frame.pack(fill=tk.BOTH, expand=True)

    def generate_custom_catalog(self):
        for widget in self.custom_list_frame.winfo_children(): widget.destroy()
        content = self.text_custom.get("1.0", tk.END).strip()
        if not content: return
        items = content.split('\n')
        self.build_catalog_ui(self.custom_list_frame, [item.strip() for item in items if item.strip()], "custom")
        self.log("✅ 自定义目录已生成。")

    def add_files(self, unique_key, category_name):
        # 更新支持的文件过滤器
        files = filedialog.askopenfilenames(
            title=f"为【{category_name}】选择文件",
            filetypes=[("支持的文件", SUPPORTED_EXTENSIONS)]
        )
        if files:
            for f in files:
                self.files_data[unique_key].append(f)
                self.listbox_dict[unique_key].insert(tk.END, os.path.basename(f))
            self.update_stats()
            self.log(f"📥 为 [{category_name}] 添加了 {len(files)} 个文件。")

    def move_item(self, unique_key, direction):
        listbox = self.listbox_dict[unique_key]
        selected_idx = listbox.curselection()
        if not selected_idx: return
        idx = selected_idx[0]
        new_idx = idx + direction
        if 0 <= new_idx < listbox.size():
            item_text = listbox.get(idx)
            listbox.delete(idx)
            listbox.insert(new_idx, item_text)
            listbox.selection_set(new_idx)
            self.files_data[unique_key][idx], self.files_data[unique_key][new_idx] = \
                self.files_data[unique_key][new_idx], self.files_data[unique_key][idx]
            self.on_listbox_select(None, listbox)

    def delete_item(self, unique_key):
        listbox = self.listbox_dict[unique_key]
        selected_idx = listbox.curselection()
        if not selected_idx: return
        idx = selected_idx[0]
        listbox.delete(idx)
        self.files_data[unique_key].pop(idx)
        self.update_stats()
        self.preview_canvas.delete("all")
        self.preview_canvas.create_text(200, 150, text="文件已移除", fill="gray", font=("微软雅黑", 10))

    def on_listbox_select(self, event, listbox=None):
        lb = listbox if listbox else event.widget
        selection = lb.curselection()
        if not selection: return
        idx = selection[0]
        lb_id = str(lb)
        if lb_id not in self.reverse_listbox_map: return
        unique_key = self.reverse_listbox_map[lb_id]
        file_path = self.files_data[unique_key][idx]
        self.render_preview(file_path)

    def render_preview(self, file_path):
        self.preview_canvas.delete("all")
        ext = os.path.splitext(file_path)[1].lower()
        try:
            # 加入更多图片格式的预览支持
            if ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff']:
                img = Image.open(file_path)
                img.thumbnail((380, 280))
                self.current_preview_img = ImageTk.PhotoImage(img)
                self.preview_canvas.create_image(200, 150, image=self.current_preview_img)
            elif ext == '.pdf':
                doc = fitz.open(file_path)
                page = doc[0]
                pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img.thumbnail((380, 280))
                self.current_preview_img = ImageTk.PhotoImage(img)
                self.preview_canvas.create_image(200, 150, image=self.current_preview_img)
                doc.close()
            elif ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.rtf']:
                self.preview_canvas.create_text(200, 150, text="📄\nOffice/文本 文档暂不支持实时预览\n将在最终合并时处理", fill="#555", font=("微软雅黑", 12), justify="center")
            else:
                self.preview_canvas.create_text(200, 150, text="未知文件格式", fill="red")
        except Exception as e:
            self.preview_canvas.create_text(200, 150, text=f"预览失败\n(文件可能已被移动或损坏)", fill="red", justify="center")

    # ================= 核心处理逻辑 =================
    def start_processing(self):
        current_tab_idx = self.notebook.index(self.notebook.select())
        if current_tab_idx == 0:
            active_catalog = DEFAULT_CATALOGS["民商事"]
            tab_prefix = "civil"
        elif current_tab_idx == 1:
            active_catalog = DEFAULT_CATALOGS["刑事"]
            tab_prefix = "criminal"
        else:
            content = self.text_custom.get("1.0", tk.END).strip()
            active_catalog = [item.strip() for item in content.split('\n') if item.strip()]
            tab_prefix = "custom"

        has_files = any(len(self.files_data.get(f"{tab_prefix}_{cat}", [])) > 0 for cat in active_catalog)
        if not has_files:
            messagebox.showwarning("空目录", "请至少在一个目录项下添加文件后再生成！")
            return

        watermark_text = simpledialog.askstring("防伪水印设置", "请输入背景水印文本\n（如果不需要水印，请留空或直接点击“取消”）")
        compress_mode = self.compress_var.get()

        save_path = filedialog.asksaveasfilename(
            title="选择归档文件保存位置", defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")], initialfile="新案件归档.pdf"
        )
        if not save_path: return

        self.btn_generate.config(state=tk.DISABLED, text="处理中，请勿操作...")
        self.txt_log.config(state=tk.NORMAL); self.txt_log.delete('1.0', tk.END); self.txt_log.config(state=tk.DISABLED)
        self.log(f"🚀 开始执行合并归档任务... (瘦身压缩模式: {'已开启' if compress_mode else '未开启'})")
        self.root.update()

        try:
            self.process_and_merge(active_catalog, save_path, tab_prefix, watermark_text, compress_mode)
            self.log("🎉 任务全部完成！")
            messagebox.showinfo("大功告成", f"归档文件已成功生成并保存至：\n{save_path}")
        except Exception as e:
            self.log(f"❌ 发生致命错误: {str(e)}")
            messagebox.showerror("发生错误", f"处理过程中发生错误：\n{str(e)}")
        finally:
            self.btn_generate.config(state=tk.NORMAL, text="🚀 一键转换并合并归档")

    def process_and_merge(self, catalog_list, save_path, tab_prefix, watermark_text, compress_mode):
        word_app, excel_app, ppt_app = None, None, None
        if IS_WINDOWS:
            try: word_app = win32com.client.DispatchEx("Word.Application"); word_app.Visible = False
            except: self.log("⚠️ 无法唤醒 Word 进程。")
            try: excel_app = win32com.client.DispatchEx("Excel.Application"); excel_app.Visible = False
            except: self.log("⚠️ 无法唤醒 Excel 进程。")
            try: ppt_app = win32com.client.DispatchEx("PowerPoint.Application") # PPT 启动特性不同，不直接设 Visible
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
                    
                    # --- 图像格式支持扩展 (.bmp, .tif, .tiff) ---
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
                    
                    # --- 纯文本/富文本支持扩展 (.txt, .rtf) ---
                    elif ext in ['.doc', '.docx', '.txt', '.rtf']:
                        if IS_WINDOWS and word_app:
                            self.log(f"   ⚙️ 调用 Word 转换: {os.path.basename(file_path)}")
                            doc = word_app.Documents.Open(os.path.abspath(file_path))
                            doc.SaveAs(os.path.abspath(temp_pdf_path), FileFormat=17) # 17 为 PDF 格式
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

                    # --- PPT 支持扩展 (.ppt, .pptx) ---
                    elif ext in ['.ppt', '.pptx']:
                        if IS_WINDOWS and ppt_app:
                            self.log(f"   ⚙️ 调用 PowerPoint 转换: {os.path.basename(file_path)}")
                            # WithWindow=False 实现后台静默打开
                            presentation = ppt_app.Presentations.Open(os.path.abspath(file_path), WithWindow=False)
                            presentation.SaveAs(os.path.abspath(temp_pdf_path), 32) # 32 为 ppSaveAsPDF
                            presentation.Close()
                            pdf_doc = fitz.open(temp_pdf_path)
                            main_pdf.insert_pdf(pdf_doc)
                            pdf_doc.close()
                        else:
                            self._create_mock_pdf(temp_pdf_path, os.path.basename(file_path))
                            doc = fitz.open(temp_pdf_path)
                            main_pdf.insert_pdf(doc)
                            doc.close()

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
                    outline.append([1, name, actual_target_page])
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
    root = tk.Tk()
    app = ArchiveApp(root)
    root.mainloop()