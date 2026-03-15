# 依赖安装: 
# 1. 基础依赖: pip install python-docx PyPDF2 python-pptx pandas tabulate openpyxl tkinterdnd2 pdfplumber
# 2. 支持 .doc 和 .ppt (旧格式) 需要安装: pip install pywin32 (仅限 Windows 系统，且需安装 Microsoft Office/WPS)

import os
import sys
import threading
import platform
from typing import List, Optional

# --- 导入第三方库 ---
# GUI
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from tkinter.constants import END

# 拖拽
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_SUPPORT = True
except ImportError:
    DND_SUPPORT = False
    # 构建 dummy 类以避免报错，实际运行时拖拽功能不可用
    class TkinterDnD:
        @staticmethod
        def Tk(): return tk.Tk()

# Word
try:
    from docx import Document
    DOCX_SUPPORT = True
except ImportError:
    DOCX_SUPPORT = False

# PDF (优先使用 pdfplumber，其次 PyPDF2)
try:
    import pdfplumber
    PDFPLUMBER_SUPPORT = True
except ImportError:
    PDFPLUMBER_SUPPORT = False

try:
    from PyPDF2 import PdfReader
    PYPDF2_SUPPORT = True
except ImportError:
    PYPDF2_SUPPORT = False

# PPT
try:
    from pptx import Presentation
    PPTX_SUPPORT = True
except ImportError:
    PPTX_SUPPORT = False

# Excel
try:
    import pandas as pd
    import tabulate
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False

# Windows COM
IS_WINDOWS = platform.system() == 'Windows'
WIN32COM_SUPPORT = False
if IS_WINDOWS:
    try:
        import win32com.client
        WIN32COM_SUPPORT = True
    except ImportError:
        pass

class FileToMarkdownConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("通用文件转 Markdown 工具 v6.0 (修复表格重复)")
        self.root.geometry("1100x750")
        
        # 获取脚本所在目录
        if getattr(sys, 'frozen', False):
            self.script_dir = os.path.dirname(sys.executable)
        else:
            self.script_dir = os.path.dirname(os.path.abspath(__file__))
            
        self.output_base_dir = os.path.join(self.script_dir, "GLM-md")
        
        self.supported_ext = {
            '.doc', '.docx', '.pdf', '.ppt', '.pptx', '.rtf',
            '.txt', '.md', '.py', '.js', '.java', '.c', '.cpp', '.h', '.hpp', 
            '.cs', '.go', '.rs', '.rb', '.php', '.swift', '.kt', '.scala', 
            '.html', '.css', '.scss', '.xml', '.json', '.yaml', '.yml', 
            '.ini', '.cfg', '.conf', '.sh', '.bat', '.cmd', '.log', '.sql',
            '.csv', '.xlsx', '.xls'
        }
        
        self.format_info = {
            "文档类": ".doc, .docx, .pdf, .ppt, .pptx, .rtf",
            "代码类": ".py, .js, .java, .c, .cpp, .go, .html, .css, .sql ...",
            "数据类": ".csv, .xlsx, .xls, .json, .xml, .yaml",
            "文本类": ".txt, .md, .log, .ini, .cfg, .sh ..."
        }
        
        self.task_list = []
        self.is_converting = False
        
        self.setup_ui()

    def setup_ui(self):
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # --- 左侧 ---
        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 工具栏
        toolbar = tk.Frame(left_frame, pady=5)
        toolbar.pack(fill=tk.X)
        
        btn_select_files = tk.Button(toolbar, text="选择文件", command=self.select_files, width=12, bg="#e1e1e1")
        btn_select_files.pack(side=tk.LEFT, padx=2)
        
        btn_select_dir = tk.Button(toolbar, text="选择文件夹", command=self.select_folder, width=12, bg="#e1e1e1")
        btn_select_dir.pack(side=tk.LEFT, padx=2)
        
        self.btn_convert = tk.Button(toolbar, text="开始转换", command=self.start_conversion_thread, width=12, bg="#4CAF50", fg="white")
        self.btn_convert.pack(side=tk.LEFT, padx=2)
        
        self.status_label = tk.Label(toolbar, text="就绪", fg="grey")
        self.status_label.pack(side=tk.RIGHT, padx=5)

        # 列表管理
        btn_frame_2 = tk.Frame(left_frame)
        btn_frame_2.pack(fill=tk.X, pady=5)
        
        btn_remove = tk.Button(btn_frame_2, text="移除选中项", command=self.remove_selected_items, width=12, bg="#ffcccc")
        btn_remove.pack(side=tk.LEFT, padx=2)
        
        btn_clear_log = tk.Button(btn_frame_2, text="清空日志", command=self.clear_log, width=12)
        btn_clear_log.pack(side=tk.LEFT, padx=2)

        # 列表区域
        list_label_frame = tk.Frame(left_frame)
        list_label_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(list_label_frame, text="待转换列表 (支持拖拽):").pack(anchor=tk.W)
        
        self.listbox = tk.Listbox(list_label_frame, selectmode=tk.EXTENDED, bg="#f9f9f9")
        self.listbox.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        scrollbar = tk.Scrollbar(list_label_frame, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)
        
        if DND_SUPPORT:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind('<<Drop>>', self.on_drop)
            self.listbox.dnd_bind('<<DragEnter>>', lambda e: self.listbox.config(bg="#d1e7dd"))
            self.listbox.dnd_bind('<<DragLeave>>', lambda e: self.listbox.config(bg="#f9f9f9"))
        
        # --- 右侧：格式说明区 ---
        right_frame = tk.LabelFrame(main_frame, text="支持的格式", padx=10, pady=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(10, 0), ipadx=10)
        
        self.format_display = tk.Text(right_frame, width=30, height=15, wrap=tk.WORD, bg="#f0f0f0", relief=tk.FLAT)
        self.format_display.pack(fill=tk.BOTH, expand=True)
        self.format_display.config(state='normal')
        
        for category, exts in self.format_info.items():
            self.format_display.insert(END, f"【{category}】\n{exts}\n\n")
        
        self.format_display.insert(END, "【特殊说明】\n")
        if PDFPLUMBER_SUPPORT:
            self.format_display.insert(END, "PDF引擎: pdfplumber (高精度)\n", "green")
        elif PYPDF2_SUPPORT:
            self.format_display.insert(END, "PDF引擎: PyPDF2 (基础)\n", "orange")
        else:
            self.format_display.insert(END, "PDF引擎: 未安装\n", "red")

        if WIN32COM_SUPPORT:
            self.format_display.insert(END, "Office组件: 已启用 (支持.doc/.ppt)\n", "green")
        else:
            self.format_display.insert(END, "Office组件: 未安装 (仅支持.docx/.pptx)\n", "red")
            
        self.format_display.config(state='disabled')
        self.format_display.tag_config("green", foreground="green")
        self.format_display.tag_config("red", foreground="red")
        self.format_display.tag_config("orange", foreground="orange")

        # --- 底部日志 ---
        log_frame = tk.Frame(left_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10,0))
        
        path_info = tk.Label(log_frame, text=f"输出目录: {self.output_base_dir}", fg="blue", anchor=tk.W)
        path_info.pack(fill=tk.X)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', bg="#f0f0f0")
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def log(self, message: str):
        self.root.after(0, self._log_safe, message)

    def _log_safe(self, message: str):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def clear_log(self):
        self.log_area.config(state='normal')
        self.log_area.delete(1.0, END)
        self.log_area.config(state='disabled')
        self.log("日志已清空。")

    def update_status(self, text: str):
        self.root.after(0, self.status_label.config, {'text': text})

    def _add_paths_to_queue(self, paths: List[str], base_source_dir: str = None):
        added_count = 0
        for path in paths:
            path = path.strip('{}')
            if not os.path.exists(path): continue

            if os.path.isfile(path):
                _, ext = os.path.splitext(path)
                if ext.lower() in self.supported_ext:
                    if base_source_dir:
                        rel_path = os.path.relpath(path, base_source_dir)
                    else:
                        rel_path = os.path.basename(path)
                    
                    if path not in [t['src'] for t in self.task_list]:
                        self.task_list.append({'src': path, 'dst_rel': rel_path})
                        added_count += 1
                else:
                    self.log(f"跳过不支持的格式: {os.path.basename(path)}")

            elif os.path.isdir(path):
                current_base = path
                self.log(f"正在扫描文件夹: {path} ...")
                for root, dirs, files in os.walk(path):
                    for file in files:
                        full_path = os.path.join(root, file)
                        _, ext = os.path.splitext(full_path)
                        if ext.lower() in self.supported_ext:
                            rel_path = os.path.relpath(full_path, current_base)
                            if full_path not in [t['src'] for t in self.task_list]:
                                self.task_list.append({'src': full_path, 'dst_rel': rel_path, 'base': current_base})
                                added_count += 1

        self.refresh_listbox()
        return added_count

    def refresh_listbox(self):
        self.listbox.delete(0, END)
        for task in self.task_list:
            display_text = f"{os.path.basename(task['src'])}  <-  {task['src']}"
            self.listbox.insert(END, display_text)

    def select_files(self):
        files = filedialog.askopenfilenames(title="选择要转换的文件")
        if files:
            count = self._add_paths_to_queue(files)
            self.log(f"添加了 {count} 个文件。")

    def select_folder(self):
        folder = filedialog.askdirectory(title="选择要转换的文件夹")
        if folder:
            count = self._add_paths_to_queue([folder], base_source_dir=folder)
            self.log(f"从文件夹添加了 {count} 个文件。")
            self.update_status(f"已添加: {os.path.basename(folder)}")

    def on_drop(self, event):
        # 解析拖拽路径
        raw_data = event.data
        paths = []
        if '{' in raw_data:
            # 处理带花括号的路径 (通常包含空格)
            temp = ""
            in_brace = False
            for char in raw_data:
                if char == '{': 
                    in_brace = True
                elif char == '}': 
                    in_brace = False
                    paths.append(temp)
                    temp = ""
                elif in_brace: 
                    temp += char
                elif char == ' ' and not in_brace:
                    if temp: 
                        paths.append(temp)
                        temp = ""
                else: 
                    temp += char
            if temp: paths.append(temp)
        else:
            paths = raw_data.split()

        if paths:
            count = 0
            for p in paths:
                if os.path.isdir(p):
                    count += self._add_paths_to_queue([p], base_source_dir=p)
                else:
                    count += self._add_paths_to_queue([p], base_source_dir=None)
            
            self.log(f"拖拽添加了 {count} 个项目。")
            self.listbox.config(bg="#f9f9f9")

    def remove_selected_items(self):
        selection = self.listbox.curselection()
        if not selection: return
        indices = sorted(selection, reverse=True)
        for i in indices:
            del self.task_list[i]
            self.listbox.delete(i)
        self.log(f"已移除 {len(indices)} 个项。")

    def start_conversion_thread(self):
        if self.is_converting:
            messagebox.showwarning("提示", "当前有任务正在执行，请稍候...")
            return
            
        if not self.task_list:
            messagebox.showwarning("提示", "列表为空，请先添加文件！")
            return
        
        self.is_converting = True
        self.btn_convert.config(state=tk.DISABLED)
        self.update_status("转换中...")
        
        # 启动后台线程
        thread = threading.Thread(target=self.convert_files_thread, daemon=True)
        thread.start()

    def convert_files_thread(self):
        try:
            if not os.path.exists(self.output_base_dir):
                os.makedirs(self.output_base_dir)
                self.log(f"创建输出根目录: {self.output_base_dir}")
        except Exception as e:
            self.log(f"错误: 无法创建输出文件夹 - {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"无法创建输出文件夹:\n{e}"))
            self.on_conversion_finished()
            return

        success_count = 0
        fail_count = 0
        total = len(self.task_list)
        
        # COM 对象复用
        word_app = None
        ppt_app = None
        
        for idx, task in enumerate(self.task_list):
            src_path = task['src']
            dst_rel = task['dst_rel']
            
            file_name = os.path.basename(src_path)
            self.update_status(f"处理中 ({idx+1}/{total}): {file_name}")
            self.log(f"正在转换: {src_path}")

            try:
                md_content = self.process_single_file(src_path, word_app, ppt_app)
                
                if md_content is not None:
                    dst_rel_md = os.path.splitext(dst_rel)[0] + ".md"
                    final_save_path = os.path.join(self.output_base_dir, dst_rel_md)
                    
                    final_dir = os.path.dirname(final_save_path)
                    if not os.path.exists(final_dir):
                        os.makedirs(final_dir)
                        
                    with open(final_save_path, 'w', encoding='utf-8') as f:
                        f.write(md_content)
                    
                    self.log(f"成功 -> {final_save_path}")
                    success_count += 1
                else:
                    fail_count += 1
                        
            except Exception as e:
                import traceback
                self.log(f"错误: {file_name} - {str(e)}")
                self.log(traceback.format_exc()) # 打印详细错误供调试
                fail_count += 1

        # 清理 COM
        if word_app: word_app.Quit()
        if ppt_app: ppt_app.Quit()

        self.log("-" * 30)
        self.log(f"转换完成！成功: {success_count}, 失败/跳过: {fail_count}")
        self.root.after(0, lambda: messagebox.showinfo("完成", f"转换完成！\n成功: {success_count}\n失败/跳过: {fail_count}\n\n文件已保存至:\n{self.output_base_dir}"))
        
        self.on_conversion_finished()

    def on_conversion_finished(self):
        self.is_converting = False
        self.root.after(0, lambda: self.btn_convert.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.update_status("就绪"))

    # ============== 文件解析核心逻辑 (优化版) ==============

    def process_single_file(self, file_path: str, word_app=None, ppt_app=None) -> Optional[str]:
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        if ext not in self.supported_ext: return None

        if ext == '.doc': return self.convert_doc_win32(file_path, word_app)
        elif ext == '.ppt': return self.convert_ppt_win32(file_path, ppt_app)
        elif ext == '.rtf': return self.convert_rtf_win32(file_path, word_app)

        if ext == '.docx': return self.convert_docx_optimized(file_path)
        elif ext == '.pdf': return self.convert_pdf_optimized(file_path)
        elif ext == '.pptx': return self.convert_pptx(file_path)
        elif ext in ('.xlsx', '.xls'): return self.convert_excel(file_path)
        elif ext == '.csv': return self.convert_csv(file_path)
        else: return self.convert_text_or_code(file_path, ext)

    # --- Word 处理 (优化：彻底解决表格重复和合并单元格问题) ---
    def convert_docx_optimized(self, file_path: str) -> str:
        if not DOCX_SUPPORT: raise Exception("缺少 python-docx 库")
        doc = Document(file_path)
        md_content = f"# {os.path.basename(file_path)}\n\n"
        
        # 1. 遍历文档元素，区分段落和表格
        # python-docx 没有 xmlDocElement，我们通过遍历 document.element.body 来保持顺序
        # 或者在遍历段落时判断其是否属于表格
        
        # 策略：遍历 document.element.body 的子元素
        from docx.oxml.ns import qn
        
        body = doc.element.body
        # 使用迭代器按顺序处理
        # 获取所有段落和表格的映射关系
        # 注意：doc.paragraphs 包含了表格内的段落，所以不能直接遍历 doc.paragraphs
        
        # 获取文档中所有顶级元素
        parent_elm = doc.element.body
        
        for child in parent_elm.iterchildren():
            if child.tag == qn('w:p'): # 段落
                # 查找对应的 Paragraph 对象
                # 简单方法：直接提取文本
                # 为了保险，我们手动解析 text
                texts = []
                for t in child.iter(qn('w:t')):
                    if t.text:
                        texts.append(t.text)
                text = "".join(texts).strip()
                
                # 尝试获取样式
                pPr = child.find(qn('w:pPr'))
                style_val = None
                if pPr is not None:
                    pStyle = pPr.find(qn('w:pStyle'))
                    if pStyle is not None:
                        style_val = pStyle.get(qn('w:val'))
                
                if text:
                    if style_val:
                        if 'Heading1' in style_val or style_val == '1':
                            md_content += f"# {text}\n\n"
                        elif 'Heading2' in style_val or style_val == '2':
                            md_content += f"## {text}\n\n"
                        elif 'Heading3' in style_val or style_val == '3':
                            md_content += f"### {text}\n\n"
                        else:
                            md_content += f"{text}\n\n"
                    else:
                        md_content += f"{text}\n\n"
                        
            elif child.tag == qn('w:tbl'): # 表格
                # 查找对应的 Table 对象。
                # 由于我们是在遍历 XML，我们需要将 XML 元素映射回 python-docx 对象，或者直接解析 XML。
                # 直接解析 XML 对合并单元格最准确，但比较繁琐。
                # 为了利用 python-docx 的 table.rows 等高级接口，我们尝试匹配对象。
                # 这里使用一个技巧：doc.tables 列表是按顺序排列的。
                pass # 下面一段逻辑将处理表格

        # 由于上述 XML 遍历匹配 Table 对象较复杂，我们采用混合策略：
        # python-docx 的 doc.tables 仅包含顶级表格。
        # 我们需要确定表格在文档中的位置。
        
        # 更简单且健壮的方案：
        # 重置内容，重新按顺序生成
        md_content = f"# {os.path.basename(file_path)}\n\n"
        
        # 获取所有块级元素
        from docx.oxml.ns import qn
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        
        # 辅助函数：判断段落是否在表格内
        def is_inside_table(paragraph):
            return paragraph._element.getparent().tag == qn('w:tc') # w:tc is table cell

        # 构建一个文档结构列表，用于按顺序输出
        # 注意：迭代 doc.element.body 是最可靠的顺序源
        # 我们将 doc.paragraphs 和 doc.tables 映射到 XML 元素
        
        # 建立 XML 到 Python 对象的映射
        para_map = {p._element: p for p in doc.paragraphs}
        table_map = {t._tbl: t for t in doc.tables}
        
        for element in doc.element.body.iterchildren():
            if element in para_map:
                p = para_map[element]
                # 必须过滤掉表格内的段落，否则表格内容会重复出现在表格前/后
                if is_inside_table(p):
                    continue
                    
                text = p.text.strip()
                if not text: continue
                
                style_name = p.style.name.lower() if p.style else ""
                
                # 处理标题
                if 'heading 1' in style_name or 'heading1' in style_name: 
                    md_content += f"# {text}\n\n"
                elif 'heading 2' in style_name or 'heading2' in style_name: 
                    md_content += f"## {text}\n\n"
                elif 'heading 3' in style_name or 'heading3' in style_name: 
                    md_content += f"### {text}\n\n"
                else: 
                    md_content += f"{text}\n\n"
                    
            elif element in table_map:
                table = table_map[element]
                md_content += self._convert_table_to_md(table)
                md_content += "\n"

        return md_content

    def _convert_table_to_md(self, table) -> str:
        """
        将 python-docx Table 对象转换为 Markdown 字符串。
        核心优化：处理合并单元格（水平和垂直），防止重复内容。
        """
        md_content = ""
        
        # 1. 扫描表格，建立二维网格，识别合并区域
        # row_count = len(table.rows) # 某些情况下可能不准确
        # col_count = len(table.columns) # 某些情况下可能不准确
        
        # 使用底层元素获取真实的网格大小
        # 获取行数
        rows = table.rows
        if not rows: return ""
        
        # 获取列数 (处理 colspan=0 的情况)
        # 检查第一行，计算实际列数
        # 但更稳妥的方法是遍历所有行找最大宽度
        grid = [] # 存储每个单元格的文本和跨度信息
        
        # 记录已经被合并的单元格坐标，避免重复输出
        # 但 python-docx 读取合并单元格时，通常会返回合并区域的左上角单元格对象
        
        # 修正逻辑：
        # 遍历行 i, 列 j。
        # 检查合并：
        # td = table.cell(i, j)
        # 检查该单元格是否已经被上方的行合并（被占位）
        
        # 复杂度较高，我们使用一种更直观的方法：
        # 使用 table._tc (table cell element) 来判断合并属性
        
        # 获取网格尺寸
        # 这里的行数是物理行数
        num_rows = len(table.rows)
        if num_rows == 0: return ""
        
        # 动态计算列数：取最大列数
        num_cols = 0
        for row in rows:
            # 计算 row 的实际网格宽度
            # row.cells 列表包含了合并后的对象引用，不能用来数数
            # 需要 XML 计算
            cols_in_row = 0
            for cell in row.cells:
                # 这是一个估算，实际上 row.cells 展平了逻辑
                pass
            # 真正的列数可以通过 table.column_cells(0) 的长度来确定吗？不行
        
        # 暴力但准确的 XML 解析法
        from docx.oxml.ns import qn
        root = table._tbl
        
        # 结果矩阵
        final_rows_text = []
        
        # 遍历每一行
        for i, tr in enumerate(root.iter(qn('w:tr'))):
            row_data = []
            # 当前行的 XML cell 元素
            tds = list(tr.findall(qn('w:tc')))
            
            # 我们需要在这一行建立一个索引，跳过被 colspan 占用的位置
            # 但 XML 中 w:gridSpan 表示水平合并，w:vMerge 表示垂直合并
            
            # 其实 python-docx 的 table.cell(i, j) 已经处理了映射
            # 问题出在：table.cell(i, j) 获取的是“逻辑单元格”
            # 如果 (i, j) 被 (i-1, j) 合并了，table.cell(i, j) 会返回上方的单元格对象
            # 这导致了用户看到的“重复”或“错乱”
            
            # 解决方案：
            # 我们需要知道逻辑上的 全局索引
            # 或者，我们直接遍历 XML，这是最干净的
            
            # XML 解析实现
            # gridSpan: 水平合并
            # vMerge: 垂直合并
            pass
            
        # --- 鉴于XML解析代码量大，采用改进的 python-docx 对象对比法 ---
        # 只要我们检测到 当前单元格 是 上面单元格的“延续”，就输出空或跳过
        
        try:
            # 获取实际的网格宽
            # 遍历第一行 XML 确认宽度？不，某些行可能不同。
            # 使用 table.columns 长度作为基准 (通常是正确的)
            if len(table.columns) > 0:
                col_count = len(table.columns)
            else:
                # 备用：计算第一行 cells
                col_count = 0
                if len(table.rows) > 0:
                     col_count = len(table.rows[0].cells) # 不准确，但作为fallback
                     
            if col_count == 0: return ""
            
            temp_content = []
            
            # 我们需要追踪哪些 cell 对象已经输出过内容了
            # 但垂直合并时，同一块区域会被看到多次
            # 如果当前 cell 在上一行出现过，且是同一对象，则是 vertical merge
            
            for r_idx in range(len(table.rows)):
                cols_in_row = []
                for c_idx in range(col_count):
                    try:
                        cell = table.cell(r_idx, c_idx)
                    except Exception:
                        cols_in_row.append(" ") # 越界
                        continue
                    
                    # 检查垂直合并
                    is_v_merge = False
                    if r_idx > 0:
                        try:
                            top_cell = table.cell(r_idx - 1, c_idx)
                            # 核心判断：如果当前单元格对象和上方单元格对象是同一个，说明是垂直合并的延续
                            if cell is top_cell:
                                is_v_merge = True
                        except:
                            pass
                            
                    # 检查水平合并
                    is_h_merge = False
                    if c_idx > 0:
                        try:
                            left_cell = table.cell(r_idx, c_idx - 1)
                            if cell is left_cell:
                                is_h_merge = True
                        except:
                            pass
                    
                    if is_v_merge or is_h_merge:
                        # 合并的延续部分，输出空占位
                        cols_in_row.append(" ")
                    else:
                        # 独立的单元格或合并的起点
                        text = cell.text.strip().replace('\n', '<br>')
                        cols_in_row.append(text if text else " ")
                        
                temp_content.append(cols_in_row)
                
            # 构建 Markdown
            for r_idx, row_vals in enumerate(temp_content):
                md_content += "| " + " | ".join(row_vals) + " |\n"
                if r_idx == 0:
                    md_content += "| " + " | ".join(["---"] * len(row_vals)) + " |\n"
                    
        except Exception as e:
            # 降级方案：简单粗暴地遍历 row.cells，虽然会有合并问题，但至少不会崩
            md_content += "\n[表格解析遇到复杂结构，已降级处理]\n"
            for row in table.rows:
                cells_text = [cell.text.strip().replace('\n', '<br>') for cell in row.cells]
                md_content += "| " + " | ".join(cells_text) + " |\n"
                
        return md_content

    # --- PDF 处理 (优化：优先使用 pdfplumber，解决格式混乱) ---
    def convert_pdf_optimized(self, file_path: str) -> str:
        md_content = f"# {os.path.basename(file_path)}\n\n"
        
        if PDFPLUMBER_SUPPORT:
            try:
                with pdfplumber.open(file_path) as pdf:
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text()
                        if text:
                            md_content += f"## Page {i+1}\n\n{text}\n\n"
                        # 可选：提取表格
                        tables = page.extract_tables()
                        for table in tables:
                            for row in table:
                                # 过滤 None
                                clean_row = [str(cell).strip().replace('\n', ' ') if cell else " " for cell in row]
                                md_content += "| " + " | ".join(clean_row) + " |\n"
                            md_content += "\n"
                return md_content
            except Exception as e:
                print(f"pdfplumber failed: {e}, falling back to PyPDF2")

        if PYPDF2_SUPPORT:
            reader = PdfReader(file_path)
            for i, page in enumerate(reader.pages):
                text = page.extract_text()
                if text: md_content += f"## Page {i+1}\n\n{text}\n\n"
            return md_content
            
        raise Exception("无可用的 PDF 解析库")

    # --- PPT 处理 ---
    def convert_pptx(self, file_path: str) -> str:
        if not PPTX_SUPPORT: raise Exception("缺少 python-pptx 库")
        prs = Presentation(file_path)
        md_content = f"# {os.path.basename(file_path)}\n\n"
        for i, slide in enumerate(prs.slides):
            md_content += f"---\n## 幻灯片 {i+1}\n\n"
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text = shape.text.strip()
                    if text: md_content += f"{text}\n\n"
        return md_content

    # --- Excel 处理 ---
    def convert_excel(self, file_path: str) -> str:
        if not EXCEL_SUPPORT: raise Exception("缺少 pandas 库")
        xlsx = pd.ExcelFile(file_path)
        md_content = f"# {os.path.basename(file_path)}\n\n"
        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name)
            md_table = tabulate.tabulate(df, headers='keys', tablefmt='pipe', showindex=False)
            md_content += f"## Sheet: {sheet_name}\n\n{md_table}\n\n"
        return md_content

    def convert_csv(self, file_path: str) -> str:
        if not EXCEL_SUPPORT:
            try:
                with open(file_path, 'r', encoding='utf-8') as f: content = f.read()
                return f"# {os.path.basename(file_path)}\n\n```csv\n{content}\n```"
            except: raise Exception("CSV 读取失败")
        df = pd.read_csv(file_path)
        md_table = tabulate.tabulate(df, headers='keys', tablefmt='pipe', showindex=False)
        return f"# {os.path.basename(file_path)}\n\n{md_table}"

    # --- 文本处理 ---
    def convert_text_or_code(self, file_path: str, ext: str) -> str:
        content = ""
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
        for enc in encodings:
            try:
                with open(file_path, 'r', encoding=enc) as f: content = f.read(); break
            except: continue
        
        if not content: return f"# {os.path.basename(file_path)}\n\n*文件内容为空或无法解码*"
        
        header = f"# {os.path.basename(file_path)}\n\n"
        code_exts = {'.py', '.js', '.java', '.c', '.cpp', '.h', '.cs', '.go', '.rs', '.rb', 
                     '.php', '.swift', '.kt', '.scala', '.html', '.css', '.scss', '.xml', 
                     '.json', '.yaml', '.yml', '.sh', '.sql'}
        if ext in code_exts:
            lang = ext[1:] if ext != '.json' else 'json'
            return header + f"```{lang}\n{content}\n```"
        else: return header + content

    # --- Win32 COM (旧格式) ---
    def _get_word_app(self, app):
        if app: return app
        if not WIN32COM_SUPPORT: return None
        try: return win32com.client.Dispatch("Word.Application")
        except: return None

    def _get_ppt_app(self, app):
        if app: return app
        if not WIN32COM_SUPPORT: return None
        try: return win32com.client.Dispatch("PowerPoint.Application")
        except: return None

    def convert_doc_win32(self, file_path: str, app) -> str:
        if not WIN32COM_SUPPORT: raise Exception("环境不支持 .doc (需要 pywin32)")
        app = self._get_word_app(app)
        if not app: raise Exception("无法启动 Word")
        doc = None
        try:
            doc = app.Documents.Open(file_path, ReadOnly=True)
            text = doc.Content.Text
            return f"# {os.path.basename(file_path)}\n\n{text}\n"
        except Exception as e: raise e
        finally:
            if doc: doc.Close(False)

    def convert_ppt_win32(self, file_path: str, app) -> str:
        if not WIN32COM_SUPPORT: raise Exception("环境不支持 .ppt (需要 pywin32)")
        app = self._get_ppt_app(app)
        if not app: raise Exception("无法启动 PowerPoint")
        pres = None
        try:
            pres = app.Presentations.Open(file_path, ReadOnly=True, Untitled=False, WithWindow=False)
            md_content = f"# {os.path.basename(file_path)}\n\n"
            for i, slide in enumerate(pres.Slides):
                md_content += f"## Slide {i+1}\n\n"
                for shape in slide.Shapes:
                    if shape.HasTextFrame:
                        md_content += shape.TextFrame.TextRange.Text + "\n\n"
            return md_content
        except Exception as e: raise e
        finally:
            if pres: pres.Close()

    def convert_rtf_win32(self, file_path: str, app) -> str:
        if not WIN32COM_SUPPORT:
            # Fallback plain text
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f: return f.read()
            except: raise Exception("无法读取 RTF")
        
        app = self._get_word_app(app)
        if not app: raise Exception("无法启动 Word")
        doc = None
        try:
            doc = app.Documents.Open(file_path, ReadOnly=True)
            return f"# {os.path.basename(file_path)}\n\n{doc.Content.Text}\n"
        except Exception as e: raise e
        finally:
            if doc: doc.Close(False)

if __name__ == "__main__":
    if DND_SUPPORT:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        
    app = FileToMarkdownConverter(root)
    root.mainloop()
