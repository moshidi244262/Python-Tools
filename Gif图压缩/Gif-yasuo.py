# 依赖安装: pip install Pillow tkinterdnd2

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
import os
import threading
from pathlib import Path
import io
import struct
import sys

# 尝试导入拖拽库
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False
    print("警告: tkinterdnd2 未安装，拖拽功能将被禁用。使用 pip install tkinterdnd2 安装")


class GifCompressor:
    def __init__(self, root):
        self.root = root
        self.root.title("GIF 压缩工具 v1.1")
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)
        
        self.gif_files = []  # 存储 GIF 文件路径（保持顺序）
        self.gif_info = {}   # 存储 GIF 信息和压缩设置
        self.is_compressing = False  # 防止重复压缩
        
        self.setup_ui()
        
    def setup_ui(self):
        # 设置样式
        style = ttk.Style()
        style.configure("TButton", padding=5)
        style.configure("TLabelframe.Label", font=("Arial", 10, "bold"))
        
        # ============ 顶部：文件选择区域 ============
        frame_top = ttk.LabelFrame(self.root, text="📁 文件选择", padding=10)
        frame_top.pack(fill=tk.X, padx=10, pady=5)
        
        # 按钮区域
        btn_frame = ttk.Frame(frame_top)
        btn_frame.pack(fill=tk.X)
        
        btn_single = ttk.Button(btn_frame, text="📄 选择GIF文件", command=self.select_files, width=15)
        btn_single.pack(side=tk.LEFT, padx=5)
        
        btn_folder = ttk.Button(btn_frame, text="📂 选择文件夹", command=self.select_folder, width=15)
        btn_folder.pack(side=tk.LEFT, padx=5)
        
        btn_clear = ttk.Button(btn_frame, text="🗑️ 清空列表", command=self.clear_list, width=15)
        btn_clear.pack(side=tk.LEFT, padx=5)
        
        # 拖拽提示
        if HAS_DND:
            lbl_dnd = ttk.Label(btn_frame, text="💡 支持拖拽文件/文件夹到下方列表区域", 
                               foreground="#0066cc", font=("Arial", 9))
            lbl_dnd.pack(side=tk.LEFT, padx=30)
        else:
            lbl_dnd = ttk.Label(btn_frame, text="⚠️ 拖拽功能未启用（需安装 tkinterdnd2）", 
                               foreground="#cc6600", font=("Arial", 9))
            lbl_dnd.pack(side=tk.LEFT, padx=30)
        
        # ============ 中部：压缩设置区域 ============
        frame_middle = ttk.LabelFrame(self.root, text="⚙️ 压缩设置", padding=10)
        frame_middle.pack(fill=tk.X, padx=10, pady=5)
        
        # --- 全局压缩板块 ---
        frame_global = ttk.LabelFrame(frame_middle, text="全局压缩", padding=10)
        frame_global.pack(fill=tk.X, pady=5)
        
        global_inner = ttk.Frame(frame_global)
        global_inner.pack(fill=tk.X)
        
        ttk.Label(global_inner, text="压缩比率:", font=("Arial", 10)).pack(side=tk.LEFT)
        
        self.global_ratio = tk.IntVar(value=70)
        
        # 比率滑块
        self.scale_ratio = ttk.Scale(global_inner, from_=99, to=1, variable=self.global_ratio, 
                                     orient=tk.HORIZONTAL, length=300, command=self.update_ratio_label)
        self.scale_ratio.pack(side=tk.LEFT, padx=10)
        
        self.lbl_ratio = ttk.Label(global_inner, text="70%", font=("Arial", 12, "bold"), 
                                   foreground="#009900", width=6)
        self.lbl_ratio.pack(side=tk.LEFT)
        
        # 快捷按钮
        ttk.Button(global_inner, text="高质量(90%)", command=lambda: self.set_global_ratio(90), 
                   width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(global_inner, text="平衡(70%)", command=lambda: self.set_global_ratio(70), 
                   width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(global_inner, text="高压缩(40%)", command=lambda: self.set_global_ratio(40), 
                   width=12).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(global_inner, text="📝 应用全局设置", command=self.apply_global_ratio, 
                   width=15).pack(side=tk.LEFT, padx=20)
        
        # --- 自定义压缩板块 ---
        frame_custom = ttk.LabelFrame(frame_middle, text="自定义压缩（双击列表项修改单个文件压缩比率）", padding=10)
        frame_custom.pack(fill=tk.X, pady=5)
        
        custom_btn_frame = ttk.Frame(frame_custom)
        custom_btn_frame.pack(fill=tk.X)
        
        ttk.Button(custom_btn_frame, text="✏️ 编辑选中项比率", command=self.edit_selected_ratio, 
                   width=18).pack(side=tk.LEFT, padx=5)
        ttk.Button(custom_btn_frame, text="📊 批量设置选中项", command=self.batch_edit_ratio, 
                   width=18).pack(side=tk.LEFT, padx=5)
        
        # --- 压缩执行按钮 ---
        compress_frame = ttk.Frame(frame_middle)
        compress_frame.pack(fill=tk.X, pady=10)
        
        self.btn_compress = ttk.Button(compress_frame, text="🚀 开始压缩全部文件", 
                                       command=self.compress_all, width=25)
        self.btn_compress.pack(side=tk.LEFT, padx=10)
        
        self.btn_compress_selected = ttk.Button(compress_frame, text="🎯 仅压缩选中文件", 
                                                command=self.compress_selected, width=25)
        self.btn_compress_selected.pack(side=tk.LEFT, padx=10)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(compress_frame, variable=self.progress_var, 
                                            maximum=100, length=300)
        self.progress_bar.pack(side=tk.LEFT, padx=20)
        
        self.lbl_progress = ttk.Label(compress_frame, text="", font=("Arial", 9))
        self.lbl_progress.pack(side=tk.LEFT, padx=5)
        
        # ============ 下部：文件列表区域 ============
        frame_list = ttk.LabelFrame(self.root, text="📋 待压缩文件列表", padding=10)
        frame_list.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建 Treeview
        columns = ("index", "name", "path", "original_size", "resolution", "frames", 
                   "ratio", "compressed_size", "reduction", "status")
        self.tree = ttk.Treeview(frame_list, columns=columns, show="headings", 
                                 selectmode="extended", height=15)
        
        # 设置列标题和宽度
        headers = [
            ("index", "序号", 50),
            ("name", "文件名", 180),
            ("path", "路径", 220),
            ("original_size", "原始大小", 90),
            ("resolution", "分辨率", 90),
            ("frames", "帧数", 60),
            ("ratio", "压缩比率", 80),
            ("compressed_size", "压缩后大小", 100),
            ("reduction", "节省空间", 100),
            ("status", "状态", 100)
        ]
        
        for col_id, header, width in headers:
            self.tree.heading(col_id, text=header)
            self.tree.column(col_id, width=width, minwidth=50)
        
        # 滚动条
        scrollbar_y = ttk.Scrollbar(frame_list, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(frame_list, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定事件
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)
        
        # 设置拖拽
        if HAS_DND:
            self.tree.drop_target_register(DND_FILES)
            self.tree.dnd_bind('<<Drop>>', self.on_drop)
            self.tree.dnd_bind('<<DragEnter>>', lambda e: self.tree.config(background='#e6f3ff'))
            self.tree.dnd_bind('<<DragLeave>>', lambda e: self.tree.config(background='white'))
        
        # ============ 底部：统计信息 ============
        frame_bottom = ttk.Frame(self.root)
        frame_bottom.pack(fill=tk.X, padx=10, pady=5)
        
        self.lbl_stats = ttk.Label(frame_bottom, text="📊 共 0 个文件，总大小: 0 B", 
                                   font=("Arial", 10, "bold"))
        self.lbl_stats.pack(side=tk.LEFT)
        
        self.lbl_output = ttk.Label(frame_bottom, text="💾 输出目录：压缩后的文件将保存在脚本所在目录的“压缩gif图片”文件夹中", 
                                    font=("Arial", 9), foreground="#666666")
        self.lbl_output.pack(side=tk.RIGHT)
    
    def update_ratio_label(self, value=None):
        ratio = self.global_ratio.get()
        self.lbl_ratio.config(text=f"{ratio}%")
        
        # 根据比率显示不同颜色
        if ratio >= 80:
            color = "#009900"  # 绿色 - 高质量
        elif ratio >= 50:
            color = "#0066cc"  # 蓝色 - 平衡
        else:
            color = "#cc6600"  # 橙色 - 高压缩
        self.lbl_ratio.config(foreground=color)
    
    def set_global_ratio(self, value):
        self.global_ratio.set(value)
        self.update_ratio_label()
    
    def select_files(self):
        """选择单个或多个GIF文件"""
        files = filedialog.askopenfilenames(
            title="选择GIF文件",
            filetypes=[("GIF文件", "*.gif *.GIF"), ("所有文件", "*.*")]
        )
        if files:
            self.add_files(list(files))
    
    def select_folder(self):
        """选择文件夹，包含子文件夹"""
        folder = filedialog.askdirectory(title="选择包含GIF文件的文件夹")
        if folder:
            self.add_folder(folder)
    
    def add_files(self, files):
        """添加文件列表"""
        added_count = 0
        for file in files:
            if file.lower().endswith('.gif'):
                if self.add_gif_file(file):
                    added_count += 1
        
        if added_count > 0:
            self.update_stats()
    
    def add_folder(self, folder):
        """添加文件夹中的所有GIF文件"""
        added_count = 0
        try:
            for root_dir, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith('.gif'):
                        filepath = os.path.join(root_dir, file)
                        if self.add_gif_file(filepath):
                            added_count += 1
        except Exception as e:
            messagebox.showerror("错误", f"遍历文件夹时出错: {str(e)}")
        
        if added_count > 0:
            self.update_stats()
            messagebox.showinfo("完成", f"已添加 {added_count} 个GIF文件")
        else:
            messagebox.showinfo("提示", "未在选定文件夹中找到GIF文件")
    
    def add_gif_file(self, filepath):
        """添加单个GIF文件到列表"""
        try:
            filepath = os.path.abspath(filepath)
            
            # 检查是否已存在
            if filepath in self.gif_files:
                return False
            
            # 获取文件信息
            size = os.path.getsize(filepath)
            if size == 0:
                return False
            
            # 尝试读取GIF信息
            with Image.open(filepath) as img:
                width, height = img.size
                resolution = f"{width}x{height}"
                
                # 计算帧数
                frame_count = 0
                try:
                    while True:
                        frame_count += 1
                        img.seek(img.tell() + 1)
                except EOFError:
                    pass
            
            # 存储信息
            self.gif_files.append(filepath)
            self.gif_info[filepath] = {
                'original_size': size,
                'width': width,
                'height': height,
                'frames': frame_count,
                'ratio': self.global_ratio.get(),
                'compressed_size': None,
                'status': '待压缩'
            }
            
            # 添加到Treeview
            self.tree.insert("", tk.END, iid=filepath, values=(
                len(self.gif_files),
                os.path.basename(filepath),
                os.path.dirname(filepath),
                self.format_size(size),
                resolution,
                frame_count,
                f"{self.global_ratio.get()}%",
                "-",
                "-",
                "待压缩"
            ))
            
            return True
            
        except Exception as e:
            print(f"添加文件失败 {filepath}: {str(e)}")
            return False
    
    def format_size(self, size):
        """格式化文件大小"""
        if size >= 1024 * 1024:
            return f"{size / (1024*1024):.2f} MB"
        elif size >= 1024:
            return f"{size / 1024:.2f} KB"
        else:
            return f"{size} B"
    
    def apply_global_ratio(self):
        """应用全局压缩比率到所有文件"""
        ratio = self.global_ratio.get()
        
        for filepath in self.gif_files:
            self.gif_info[filepath]['ratio'] = ratio
            info = self.gif_info[filepath]
            
            self.tree.item(filepath, values=(
                self.tree.item(filepath)['values'][0],  # 序号
                os.path.basename(filepath),
                os.path.dirname(filepath),
                self.format_size(info['original_size']),
                f"{info['width']}x{info['height']}",
                info['frames'],
                f"{ratio}%",
                "-",
                "-",
                "待压缩"
            ))
        
        messagebox.showinfo("完成", f"已将所有文件压缩比率设置为 {ratio}%")
    
    def on_double_click(self, event):
        """双击编辑压缩比率"""
        self.edit_selected_ratio()
    
    def edit_selected_ratio(self):
        """编辑选中文件的压缩比率"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请先选择要编辑的文件")
            return
        
        if len(selected) == 1:
            # 单个文件编辑
            self.edit_single_ratio(selected[0])
        else:
            # 多个文件批量编辑
            self.batch_edit_ratio()
    
    def edit_single_ratio(self, item):
        """编辑单个文件的压缩比率"""
        filepath = item
        current_ratio = self.gif_info[filepath]['ratio']
        
        # 创建编辑窗口
        edit_win = tk.Toplevel(self.root)
        edit_win.title("设置压缩比率")
        edit_win.geometry("400x200")
        edit_win.resizable(False, False)
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        # 居中显示
        edit_win.geometry(f"+{self.root.winfo_x() + 350}+{self.root.winfo_y() + 250}")
        
        ttk.Label(edit_win, text=f"文件: {os.path.basename(filepath)}", 
                 font=("Arial", 10, "bold")).pack(pady=10)
        
        ttk.Label(edit_win, text=f"原始大小: {self.format_size(self.gif_info[filepath]['original_size'])}").pack()
        
        # 比率设置
        ratio_frame = ttk.Frame(edit_win)
        ratio_frame.pack(pady=15)
        
        ttk.Label(ratio_frame, text="压缩比率:").pack(side=tk.LEFT, padx=5)
        
        ratio_var = tk.IntVar(value=current_ratio)
        scale = ttk.Scale(ratio_frame, from_=99, to=1, variable=ratio_var, 
                         orient=tk.HORIZONTAL, length=200, command=lambda v: lbl.config(text=f"{int(float(v))}%"))
        scale.pack(side=tk.LEFT, padx=5)
        
        lbl = ttk.Label(ratio_frame, text=f"{current_ratio}%", font=("Arial", 11, "bold"), width=5)
        lbl.pack(side=tk.LEFT, padx=5)
        
        def apply():
            new_ratio = ratio_var.get()
            self.gif_info[filepath]['ratio'] = new_ratio
            self.tree.set(filepath, "ratio", f"{new_ratio}%")
            self.tree.set(filepath, "status", "待压缩")
            self.tree.set(filepath, "compressed_size", "-")
            self.tree.set(filepath, "reduction", "-")
            edit_win.destroy()
        
        ttk.Button(edit_win, text="✓ 应用", command=apply, width=15).pack(pady=15)
    
    def batch_edit_ratio(self):
        """批量编辑选中文件的压缩比率"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请先选择要编辑的文件")
            return
        
        # 创建编辑窗口
        edit_win = tk.Toplevel(self.root)
        edit_win.title("批量设置压缩比率")
        edit_win.geometry("400x180")
        edit_win.resizable(False, False)
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        edit_win.geometry(f"+{self.root.winfo_x() + 350}+{self.root.winfo_y() + 250}")
        
        ttk.Label(edit_win, text=f"已选择 {len(selected)} 个文件", 
                 font=("Arial", 10, "bold")).pack(pady=15)
        
        # 比率设置
        ratio_frame = ttk.Frame(edit_win)
        ratio_frame.pack(pady=10)
        
        ttk.Label(ratio_frame, text="压缩比率:").pack(side=tk.LEFT, padx=5)
        
        ratio_var = tk.IntVar(value=self.global_ratio.get())
        scale = ttk.Scale(ratio_frame, from_=99, to=1, variable=ratio_var, 
                         orient=tk.HORIZONTAL, length=200, command=lambda v: lbl.config(text=f"{int(float(v))}%"))
        scale.pack(side=tk.LEFT, padx=5)
        
        lbl = ttk.Label(ratio_frame, text=f"{self.global_ratio.get()}%", font=("Arial", 11, "bold"), width=5)
        lbl.pack(side=tk.LEFT, padx=5)
        
        def apply():
            new_ratio = ratio_var.get()
            for item in selected:
                self.gif_info[item]['ratio'] = new_ratio
                self.tree.set(item, "ratio", f"{new_ratio}%")
                self.tree.set(item, "status", "待压缩")
                self.tree.set(item, "compressed_size", "-")
                self.tree.set(item, "reduction", "-")
            edit_win.destroy()
        
        ttk.Button(edit_win, text="✓ 应用到全部选中文件", command=apply, width=20).pack(pady=15)
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="编辑压缩比率", command=self.edit_selected_ratio)
        menu.add_command(label="移除选中项", command=self.remove_selected)
        menu.add_separator()
        menu.add_command(label="清空列表", command=self.clear_list)
        menu.add_separator()
        menu.add_command(label="打开文件所在目录", command=self.open_file_location)
        
        menu.post(event.x_root, event.y_root)
    
    def remove_selected(self):
        """移除选中的文件"""
        selected = self.tree.selection()
        for item in selected:
            if item in self.gif_files:
                self.gif_files.remove(item)
            if item in self.gif_info:
                del self.gif_info[item]
            self.tree.delete(item)
        
        # 重新编号
        self.reindex_files()
        self.update_stats()
    
    def reindex_files(self):
        """重新编号文件"""
        for idx, filepath in enumerate(self.gif_files, 1):
            self.tree.set(filepath, "index", idx)
    
    def open_file_location(self):
        """打开文件所在目录"""
        selected = self.tree.selection()
        if selected:
            filepath = selected[0]
            directory = os.path.dirname(filepath)
            if os.path.exists(directory):
                os.startfile(directory)
    
    def clear_list(self):
        """清空文件列表"""
        self.gif_files.clear()
        self.gif_info.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.update_stats()
    
    def update_stats(self):
        """更新统计信息"""
        total_count = len(self.gif_files)
        total_size = sum(info['original_size'] for info in self.gif_info.values())
        
        self.lbl_stats.config(text=f"📊 共 {total_count} 个文件，总大小: {self.format_size(total_size)}")
    
    def on_drop(self, event):
        """处理拖拽事件"""
        try:
            dropped = self.root.tk.splitlist(event.data)
            added_count = 0
            
            for item in dropped:
                # 处理Windows路径中的花括号
                item = item.strip('{}')
                
                if os.path.isfile(item):
                    if item.lower().endswith('.gif'):
                        if self.add_gif_file(item):
                            added_count += 1
                elif os.path.isdir(item):
                    for root_dir, dirs, files in os.walk(item):
                        for file in files:
                            if file.lower().endswith('.gif'):
                                filepath = os.path.join(root_dir, file)
                                if self.add_gif_file(filepath):
                                    added_count += 1
            
            if added_count > 0:
                self.update_stats()
                self.reindex_files()
                
        except Exception as e:
            messagebox.showerror("错误", f"拖拽处理失败: {str(e)}")
        finally:
            self.tree.config(background='white')
    
    def compress_all(self):
        """压缩所有文件"""
        if not self.gif_files:
            messagebox.showwarning("提示", "请先添加GIF文件")
            return
        
        self.do_compress(self.gif_files)
    
    def compress_selected(self):
        """仅压缩选中的文件"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请先选择要压缩的文件")
            return
        
        self.do_compress(list(selected))
    
    def do_compress(self, files_to_compress):
        """执行压缩（在新线程中）"""
        if self.is_compressing:
            messagebox.showwarning("提示", "正在压缩中，请稍候...")
            return
        
        # 检查原文件是否存在
        for filepath in files_to_compress:
            if not os.path.exists(filepath):
                messagebox.showerror("错误", f"文件不存在: {filepath}")
                return
        
        self.is_compressing = True
        self.btn_compress.config(state=tk.DISABLED)
        self.btn_compress_selected.config(state=tk.DISABLED)
        
        # 在新线程中执行压缩
        thread = threading.Thread(target=self._compress_thread, args=(files_to_compress,))
        thread.daemon = True
        thread.start()
    
    def _compress_thread(self, files_to_compress):
        """压缩线程"""
        total_files = len(files_to_compress)
        success_count = 0
        fail_count = 0
        
        for idx, filepath in enumerate(files_to_compress, 1):
            try:
                # 更新进度
                self.root.after(0, lambda i=idx, t=total_files: 
                               self.update_progress(i, t, os.path.basename(filepath)))
                
                result = self.compress_single_gif(filepath)
                
                if result['success']:
                    success_count += 1
                    self.root.after(0, lambda f=filepath, r=result: 
                                   self.update_file_result(f, r))
                else:
                    fail_count += 1
                    self.root.after(0, lambda f=filepath, e=result.get('error', '未知错误'): 
                                   self.update_file_error(f, e))
                    
            except Exception as e:
                fail_count += 1
                self.root.after(0, lambda f=filepath, e=str(e): 
                               self.update_file_error(f, e))
        
        # 完成后更新UI
        self.root.after(0, lambda: self.compress_complete(success_count, fail_count))
    
    def update_progress(self, current, total, filename):
        """更新进度显示"""
        progress = (current - 1) / total * 100
        self.progress_var.set(progress)
        self.lbl_progress.config(text=f"({current}/{total}) {filename[:25]}...")
    
    def compress_single_gif(self, filepath):
        """压缩单个GIF文件"""
        try:
            info = self.gif_info[filepath]
            ratio = info['ratio']
            
            # 打开GIF文件
            with Image.open(filepath) as img:
                frames = []
                durations = []
                
                # 提取所有帧
                try:
                    while True:
                        # 复制当前帧
                        frame = img.copy()
                        
                        # 处理调色板
                        if frame.mode not in ('P', 'L'):
                            frame = frame.convert('P', palette=Image.ADAPTIVE, colors=256)
                        
                        frames.append(frame)
                        durations.append(img.info.get('duration', 100))
                        img.seek(img.tell() + 1)
                except EOFError:
                    pass
                
                if not frames:
                    return {'success': False, 'error': '无法读取GIF帧'}
                
                # 计算新尺寸（压缩比率越高，尺寸越大；比率越低，压缩程度越高）
                original_width = info['width']
                original_height = info['height']
                scale_factor = ratio / 100
                new_width = max(16, int(original_width * scale_factor))
                new_height = max(16, int(original_height * scale_factor))
                
                # 计算颜色数量（根据压缩比率调整）
                colors = max(16, int(256 * scale_factor))
                
                compressed_frames = []
                
                for frame in frames:
                    # 统一调色板模式
                    if frame.mode == 'P':
                        frame = frame.convert('RGB').convert('P', 
                                                             palette=Image.ADAPTIVE, 
                                                             colors=colors)
                    
                    # 缩放
                    if new_width != original_width or new_height != original_height:
                        frame = frame.resize((new_width, new_height), Image.LANCZOS)
                    
                    compressed_frames.append(frame)
                
                # ================= 新增逻辑：指定输出目录为“压缩gif图片” =================
                # 获取当前脚本所在目录
                try:
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                except NameError:
                    script_dir = os.getcwd() # 兜底逻辑，以防在特殊环境（如Jupyter）中运行
                
                output_dir = os.path.join(script_dir, "压缩gif图片")
                
                # 如果文件夹不存在则自动创建
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                    
                base_name = os.path.splitext(os.path.basename(filepath))[0]
                output_name = f"{base_name}_compressed.gif"
                output_path = os.path.join(output_dir, output_name)
                
                # 处理重复文件名（如果多次压缩同名文件，不覆盖旧的，而是添加序号）
                counter = 1
                while os.path.exists(output_path):
                    output_name = f"{base_name}_compressed_{counter}.gif"
                    output_path = os.path.join(output_dir, output_name)
                    counter += 1
                # ====================================================================
                
                # 保存压缩后的GIF
                compressed_frames[0].save(
                    output_path,
                    save_all=True,
                    append_images=compressed_frames[1:],
                    duration=durations,
                    loop=img.info.get('loop', 0),
                    optimize=True,
                    disposal=2
                )
                
                # 获取压缩后的大小
                compressed_size = os.path.getsize(output_path)
                original_size = info['original_size']
                
                # 计算节省空间
                if compressed_size < original_size:
                    reduction = (1 - compressed_size / original_size) * 100
                else:
                    reduction = -(compressed_size / original_size - 1) * 100
                
                return {
                    'success': True,
                    'output_path': output_path,
                    'compressed_size': compressed_size,
                    'reduction': reduction
                }
                
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def update_file_result(self, filepath, result):
        """更新文件压缩结果"""
        info = self.gif_info[filepath]
        info['compressed_size'] = result['compressed_size']
        info['status'] = '已完成'
        
        reduction = result['reduction']
        if reduction > 0:
            reduction_str = f"↓ {reduction:.1f}%"
            status = "✓ 已完成"
        else:
            reduction_str = f"↑ {abs(reduction):.1f}%"
            status = "⚠ 已增大"
        
        self.tree.item(filepath, values=(
            self.tree.item(filepath)['values'][0],
            os.path.basename(filepath),
            os.path.dirname(filepath),
            self.format_size(info['original_size']),
            f"{info['width']}x{info['height']}",
            info['frames'],
            f"{info['ratio']}%",
            self.format_size(result['compressed_size']),
            reduction_str,
            status
        ))
    
    def update_file_error(self, filepath, error):
        """更新文件错误状态"""
        info = self.gif_info[filepath]
        info['status'] = '失败'
        
        self.tree.item(filepath, values=(
            self.tree.item(filepath)['values'][0],
            os.path.basename(filepath),
            os.path.dirname(filepath),
            self.format_size(info['original_size']),
            f"{info['width']}x{info['height']}",
            info['frames'],
            f"{info['ratio']}%",
            "-",
            "-",
            f"✗ 失败"
        ))
    
    def compress_complete(self, success_count, fail_count):
        """压缩完成"""
        self.is_compressing = False
        self.btn_compress.config(state=tk.NORMAL)
        self.btn_compress_selected.config(state=tk.NORMAL)
        self.progress_var.set(100)
        
        total_original = sum(
            self.gif_info[f]['original_size'] 
            for f in self.gif_files 
            if self.gif_info[f].get('compressed_size')
        )
        total_compressed = sum(
            self.gif_info[f]['compressed_size'] 
            for f in self.gif_files 
            if self.gif_info[f].get('compressed_size')
        )
        
        if total_original > 0:
            saved = total_original - total_compressed
            saved_percent = (1 - total_compressed / total_original) * 100
            saved_str = self.format_size(abs(saved))
            
            if saved > 0:
                size_msg = f"，共节省 {saved_str} ({saved_percent:.1f}%)"
            else:
                size_msg = f"，总大小增加 {saved_str}"
        else:
            size_msg = ""
        
        msg = f"压缩完成！成功: {success_count} 个，失败: {fail_count} 个{size_msg}"
        messagebox.showinfo("压缩完成", msg)
        
        self.lbl_progress.config(text="完成!")


def main():
    """主函数"""
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    
    # 设置窗口图标（如果有的话）
    try:
        root.iconbitmap(default='')
    except:
        pass
    
    app = GifCompressor(root)
    root.mainloop()


if __name__ == "__main__":
    main()
