# 依赖安装: pip install Pillow tkinterdnd2

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
import os
import threading
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
        self.root.title("GIF 压缩工具 v2.0 (优化版)")
        self.root.geometry("1150x780")
        self.root.minsize(950, 650)
        
        self.gif_files = []  # 存储 GIF 文件路径（保持顺序）
        self.gif_info = {}   # 存储 GIF 信息和压缩设置
        self.is_compressing = False  # 防止重复压缩
        self.cancel_flag = False     # 取消标志位
        
        # 确定脚本或程序的运行目录 (兼容 PyInstaller 打包)
        if getattr(sys, 'frozen', False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
            
        self.default_output_dir = os.path.join(self.base_dir, "压缩gif图片")
        
        self.setup_ui()
        
    def setup_ui(self):
        # ============ UI 美化与主题设置 ============
        style = ttk.Style()
        # 尝试使用更现代的内置主题（'clam' 比默认的 Windows 主题更好看）
        if 'clam' in style.theme_names():
            style.theme_use('clam')
            
        style.configure("TButton", padding=6, font=("Microsoft YaHei", 9))
        style.configure("TLabelframe.Label", font=("Microsoft YaHei", 10, "bold"), foreground="#333333")
        style.configure("Treeview.Heading", font=("Microsoft YaHei", 9, "bold"))
        style.configure("Treeview", font=("Microsoft YaHei", 9), rowheight=25)
        
        # ============ 顶部：文件选择区域 ============
        frame_top = ttk.LabelFrame(self.root, text="📁 文件选择", padding=10)
        frame_top.pack(fill=tk.X, padx=15, pady=8)
        
        btn_frame = ttk.Frame(frame_top)
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="📄 选择GIF文件", command=self.select_files, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📂 选择文件夹", command=self.select_folder, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 清空列表", command=self.clear_list, width=15).pack(side=tk.LEFT, padx=5)
        
        if HAS_DND:
            lbl_dnd = ttk.Label(btn_frame, text="💡 提示: 支持直接拖拽文件/文件夹到下方列表区域", foreground="#0066cc")
        else:
            lbl_dnd = ttk.Label(btn_frame, text="⚠️ 拖拽未启用 (请在命令行运行: pip install tkinterdnd2)", foreground="#cc6600")
        lbl_dnd.pack(side=tk.LEFT, padx=20)
        
        # ============ 中部：压缩设置区域 ============
        frame_middle = ttk.LabelFrame(self.root, text="⚙️ 压缩设置与执行", padding=10)
        frame_middle.pack(fill=tk.X, padx=15, pady=5)
        
        global_inner = ttk.Frame(frame_middle)
        global_inner.pack(fill=tk.X, pady=5)
        
        ttk.Label(global_inner, text="全局压缩强度:").pack(side=tk.LEFT, padx=5)
        self.global_ratio = tk.IntVar(value=70)
        
        self.scale_ratio = ttk.Scale(global_inner, from_=99, to=1, variable=self.global_ratio, 
                                     orient=tk.HORIZONTAL, length=250, command=self.update_ratio_label)
        self.scale_ratio.pack(side=tk.LEFT, padx=10)
        
        self.lbl_ratio = ttk.Label(global_inner, text="70%", font=("Arial", 11, "bold"), foreground="#0066cc", width=5)
        self.lbl_ratio.pack(side=tk.LEFT)
        
        ttk.Button(global_inner, text="📝 应用到所有文件", command=self.apply_global_ratio).pack(side=tk.LEFT, padx=15)
        
        # 分隔符
        ttk.Separator(global_inner, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=15)
        
        self.btn_compress = ttk.Button(global_inner, text="🚀 开始压缩全部", command=self.compress_all)
        self.btn_compress.pack(side=tk.LEFT, padx=5)
        
        self.btn_compress_selected = ttk.Button(global_inner, text="🎯 仅压缩选中", command=self.compress_selected)
        self.btn_compress_selected.pack(side=tk.LEFT, padx=5)
        
        self.btn_cancel = ttk.Button(global_inner, text="⏹️ 取消操作", command=self.cancel_compression, state=tk.DISABLED)
        self.btn_cancel.pack(side=tk.LEFT, padx=5)
        
        # 进度条
        progress_frame = ttk.Frame(frame_middle)
        progress_frame.pack(fill=tk.X, pady=10)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        self.lbl_progress = ttk.Label(progress_frame, text="等待开始...", width=30, anchor=tk.W)
        self.lbl_progress.pack(side=tk.RIGHT)
        
        # ============ 下部：文件列表区域 ============
        frame_list = ttk.LabelFrame(self.root, text="📋 待压缩文件列表 (双击修改单个文件比率)", padding=10)
        frame_list.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        columns = ("index", "name", "path", "original_size", "resolution", "frames", 
                   "ratio", "compressed_size", "reduction", "status")
        self.tree = ttk.Treeview(frame_list, columns=columns, show="headings", selectmode="extended")
        
        headers = [
            ("index", "序号", 40), ("name", "文件名", 180), ("path", "路径", 200),
            ("original_size", "原大小", 80), ("resolution", "分辨率", 80), 
            ("frames", "帧数", 50), ("ratio", "压缩比", 60), 
            ("compressed_size", "新大小", 80), ("reduction", "节省空间", 80), 
            ("status", "状态", 80)
        ]
        
        for col_id, header, width in headers:
            self.tree.heading(col_id, text=header)
            self.tree.column(col_id, width=width, minwidth=40, anchor=tk.CENTER if col_id not in ("name", "path") else tk.W)
        
        # 设置斑马纹背景
        self.tree.tag_configure('oddrow', background='#F5F5F5')
        self.tree.tag_configure('evenrow', background='#FFFFFF')
        
        scrollbar_y = ttk.Scrollbar(frame_list, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(frame_list, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 绑定事件
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)
        
        if HAS_DND:
            self.tree.drop_target_register(DND_FILES)
            self.tree.dnd_bind('<<Drop>>', self.on_drop)
        
        # ============ 底部：输出与统计 ============
        frame_bottom = ttk.Frame(self.root)
        frame_bottom.pack(fill=tk.X, padx=15, pady=10)
        
        self.lbl_stats = ttk.Label(frame_bottom, text="📊 共 0 个文件，总大小: 0 B", font=("Microsoft YaHei", 9, "bold"))
        self.lbl_stats.pack(side=tk.LEFT)
        
        # 输出目录设置
        out_frame = ttk.Frame(frame_bottom)
        out_frame.pack(side=tk.RIGHT)
        
        ttk.Label(out_frame, text="保存至:").pack(side=tk.LEFT)
        self.output_dir_var = tk.StringVar(value=self.default_output_dir)
        ttk.Entry(out_frame, textvariable=self.output_dir_var, width=40, state="readonly").pack(side=tk.LEFT, padx=5)
        ttk.Button(out_frame, text="更改目录", command=self.change_output_dir).pack(side=tk.LEFT, padx=2)
        ttk.Button(out_frame, text="打开文件夹", command=self.open_output_dir).pack(side=tk.LEFT, padx=2)
    
    def update_ratio_label(self, value=None):
        ratio = self.global_ratio.get()
        self.lbl_ratio.config(text=f"{ratio}%")
        if ratio >= 80:
            self.lbl_ratio.config(foreground="#009900")
        elif ratio >= 50:
            self.lbl_ratio.config(foreground="#0066cc")
        else:
            self.lbl_ratio.config(foreground="#cc6600")
            
    def change_output_dir(self):
        new_dir = filedialog.askdirectory(title="选择压缩后的保存目录", initialdir=self.output_dir_var.get())
        if new_dir:
            self.output_dir_var.set(new_dir)
            
    def open_output_dir(self):
        folder = self.output_dir_var.get()
        if not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)
        os.startfile(folder)
    
    def select_files(self):
        files = filedialog.askopenfilenames(title="选择GIF文件", filetypes=[("GIF文件", "*.gif *.GIF")])
        if files:
            self.add_files(list(files))
    
    def select_folder(self):
        folder = filedialog.askdirectory(title="选择包含GIF的文件夹")
        if folder:
            self.add_folder(folder)
    
    def add_files(self, files):
        added = sum(1 for f in files if f.lower().endswith('.gif') and self.add_gif_file(f))
        if added > 0: self.refresh_tree_tags()
    
    def add_folder(self, folder):
        added = 0
        for root_dir, _, files in os.walk(folder):
            for file in files:
                if file.lower().endswith('.gif') and self.add_gif_file(os.path.join(root_dir, file)):
                    added += 1
        if added > 0:
            self.refresh_tree_tags()
            messagebox.showinfo("完成", f"已添加 {added} 个GIF文件")
    
    def add_gif_file(self, filepath):
        filepath = os.path.abspath(filepath)
        if filepath in self.gif_files or os.path.getsize(filepath) == 0:
            return False
            
        try:
            size = os.path.getsize(filepath)
            with Image.open(filepath) as img:
                width, height = img.size
                frame_count = img.n_frames if hasattr(img, 'n_frames') else 1
            
            self.gif_files.append(filepath)
            self.gif_info[filepath] = {
                'original_size': size, 'width': width, 'height': height,
                'frames': frame_count, 'ratio': self.global_ratio.get(),
                'compressed_size': None, 'status': '待压缩'
            }
            
            self.tree.insert("", tk.END, iid=filepath, values=(
                len(self.gif_files), os.path.basename(filepath), os.path.dirname(filepath),
                self.format_size(size), f"{width}x{height}", frame_count,
                f"{self.global_ratio.get()}%", "-", "-", "待压缩"
            ))
            self.update_stats()
            return True
        except Exception as e:
            print(f"添加文件失败: {e}")
            return False
            
    def refresh_tree_tags(self):
        """刷新斑马纹背景"""
        for index, item in enumerate(self.tree.get_children()):
            tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.item(item, tags=(tag,))
    
    def format_size(self, size):
        if size >= 1048576: return f"{size / 1048576:.2f} MB"
        if size >= 1024: return f"{size / 1024:.2f} KB"
        return f"{size} B"
    
    def apply_global_ratio(self):
        ratio = self.global_ratio.get()
        for filepath in self.gif_files:
            self.gif_info[filepath]['ratio'] = ratio
            # 优化：只更新需要修改的列，避免重写整行导致的问题
            self.tree.set(filepath, "ratio", f"{ratio}%")
            self.tree.set(filepath, "status", "待压缩")
            self.tree.set(filepath, "compressed_size", "-")
            self.tree.set(filepath, "reduction", "-")
        messagebox.showinfo("提示", f"已将所有文件压缩比率更新为 {ratio}%")
    
    def on_double_click(self, event):
        selected = self.tree.selection()
        if selected: self.edit_single_ratio(selected[0])
            
    def edit_single_ratio(self, filepath):
        current_ratio = self.gif_info[filepath]['ratio']
        edit_win = tk.Toplevel(self.root)
        edit_win.title("修改比率")
        edit_win.geometry("350x150")
        edit_win.transient(self.root)
        edit_win.grab_set()
        edit_win.geometry(f"+{self.root.winfo_x() + 400}+{self.root.winfo_y() + 300}")
        
        ttk.Label(edit_win, text=f"正在修改: {os.path.basename(filepath)[:20]}...").pack(pady=10)
        
        frame = ttk.Frame(edit_win)
        frame.pack(pady=10)
        ratio_var = tk.IntVar(value=current_ratio)
        ttk.Scale(frame, from_=99, to=1, variable=ratio_var, orient=tk.HORIZONTAL, length=150, 
                  command=lambda v: lbl.config(text=f"{int(float(v))}%")).pack(side=tk.LEFT)
        lbl = ttk.Label(frame, text=f"{current_ratio}%", width=5)
        lbl.pack(side=tk.LEFT, padx=10)
        
        def apply():
            nr = ratio_var.get()
            self.gif_info[filepath]['ratio'] = nr
            self.tree.set(filepath, "ratio", f"{nr}%")
            self.tree.set(filepath, "status", "待压缩")
            edit_win.destroy()
            
        ttk.Button(edit_win, text="确定", command=apply).pack()

    def show_context_menu(self, event):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="🗑️ 移除选中项", command=self.remove_selected)
        menu.add_command(label="📂 打开所在目录", command=self.open_file_location)
        menu.add_command(label="👁️ 预览原图", command=self.preview_image)
        menu.post(event.x_root, event.y_root)
        
    def preview_image(self):
        sel = self.tree.selection()
        if sel: os.startfile(sel[0])

    def remove_selected(self):
        for item in self.tree.selection():
            if item in self.gif_files: self.gif_files.remove(item)
            if item in self.gif_info: del self.gif_info[item]
            self.tree.delete(item)
        for idx, filepath in enumerate(self.gif_files, 1):
            self.tree.set(filepath, "index", idx)
        self.update_stats()
        self.refresh_tree_tags()
    
    def open_file_location(self):
        sel = self.tree.selection()
        if sel: os.startfile(os.path.dirname(sel[0]))
    
    def clear_list(self):
        self.gif_files.clear()
        self.gif_info.clear()
        for item in self.tree.get_children(): self.tree.delete(item)
        self.update_stats()
    
    def update_stats(self):
        tsize = sum(i['original_size'] for i in self.gif_info.values())
        self.lbl_stats.config(text=f"📊 共 {len(self.gif_files)} 个文件，总大小: {self.format_size(tsize)}")
    
    def on_drop(self, event):
        dropped = self.root.tk.splitlist(event.data)
        for item in dropped:
            item = item.strip('{}')
            if os.path.isfile(item): self.add_files([item])
            elif os.path.isdir(item): self.add_folder(item)
            
    def cancel_compression(self):
        """取消压缩操作"""
        if self.is_compressing:
            self.cancel_flag = True
            self.btn_cancel.config(state=tk.DISABLED, text="正在停止...")

    def compress_all(self):
        if not self.gif_files: return messagebox.showwarning("提示", "请先添加文件")
        self.do_compress(self.gif_files)
    
    def compress_selected(self):
        sel = self.tree.selection()
        if not sel: return messagebox.showwarning("提示", "请先选择文件")
        self.do_compress(list(sel))
    
    def do_compress(self, files_to_compress):
        if self.is_compressing: return
        self.is_compressing = True
        self.cancel_flag = False
        
        self.btn_compress.config(state=tk.DISABLED)
        self.btn_compress_selected.config(state=tk.DISABLED)
        self.btn_cancel.config(state=tk.NORMAL, text="⏹️ 取消操作")
        
        threading.Thread(target=self._compress_thread, args=(files_to_compress,), daemon=True).start()
    
    def _compress_thread(self, files):
        total = len(files)
        succ = fail = 0
        
        # 确保输出目录存在
        out_dir = self.output_dir_var.get()
        os.makedirs(out_dir, exist_ok=True)
        
        for idx, filepath in enumerate(files, 1):
            if self.cancel_flag:
                self.root.after(0, lambda: self.lbl_progress.config(text="用户已取消!"))
                break
                
            self.root.after(0, self.update_progress, idx, total, os.path.basename(filepath))
            result = self.compress_single_gif(filepath, out_dir)
            
            if result['success']:
                succ += 1
                self.root.after(0, self.update_file_result, filepath, result)
            else:
                fail += 1
                self.root.after(0, self.update_file_error, filepath, result.get('error', '未知错误'))
                
        self.root.after(0, self.compress_complete, succ, fail, self.cancel_flag)
    
    def update_progress(self, current, total, filename):
        self.progress_var.set((current - 1) / total * 100)
        self.lbl_progress.config(text=f"[{current}/{total}] 正在处理: {filename[:15]}...")
    
    def compress_single_gif(self, filepath, output_dir):
        try:
            info = self.gif_info[filepath]
            ratio = info['ratio'] / 100.0
            
            with Image.open(filepath) as img:
                frames, durations = [], []
                # 获取原图的循环次数，如果没有则默认为 0（无限循环）
                loop = img.info.get('loop', 0) 
                
                try:
                    while True:
                        frame = img.copy()
                        if frame.mode not in ('P', 'L'):
                            frame = frame.convert('P', palette=Image.ADAPTIVE, colors=256)
                        frames.append(frame)
                        durations.append(img.info.get('duration', 100))
                        img.seek(img.tell() + 1)
                except EOFError: pass
                
                if not frames: return {'success': False, 'error': '无法读取GIF帧'}
                
                new_w = max(16, int(info['width'] * ratio))
                new_h = max(16, int(info['height'] * ratio))
                colors = max(16, int(256 * ratio))
                
                comp_frames = []
                for f in frames:
                    if f.mode == 'P':
                        f = f.convert('RGB').convert('P', palette=Image.ADAPTIVE, colors=colors)
                    if new_w != info['width'] or new_h != info['height']:
                        f = f.resize((new_w, new_h), Image.Resampling.LANCZOS)
                    comp_frames.append(f)
                
                base_name = os.path.splitext(os.path.basename(filepath))[0]
                out_path = os.path.join(output_dir, f"{base_name}_comp.gif")
                counter = 1
                while os.path.exists(out_path):
                    out_path = os.path.join(output_dir, f"{base_name}_comp_{counter}.gif")
                    counter += 1
                
                comp_frames[0].save(
                    out_path, save_all=True, append_images=comp_frames[1:],
                    duration=durations, loop=loop, optimize=True, disposal=2
                )
                
                c_size = os.path.getsize(out_path)
                o_size = info['original_size']
                reduction = (1 - c_size / o_size) * 100 if c_size < o_size else -(c_size / o_size - 1) * 100
                
                return {'success': True, 'compressed_size': c_size, 'reduction': reduction}
        except Exception as e:
            return {'success': False, 'error': str(e)}
            
    def update_file_result(self, filepath, res):
        # 使用 set 方法单独更新列，防止破坏整行数据结构
        self.tree.set(filepath, "compressed_size", self.format_size(res['compressed_size']))
        self.gif_info[filepath]['compressed_size'] = res['compressed_size']
        
        red = res['reduction']
        if red > 0:
            self.tree.set(filepath, "reduction", f"↓ {red:.1f}%")
            self.tree.set(filepath, "status", "✅ 完成")
        else:
            self.tree.set(filepath, "reduction", f"↑ {abs(red):.1f}%")
            self.tree.set(filepath, "status", "⚠️ 变大")

    def update_file_error(self, filepath, err):
        self.tree.set(filepath, "status", "❌ 失败")
        print(f"[{filepath}] 压缩失败: {err}")

    def compress_complete(self, succ, fail, was_canceled):
        self.is_compressing = False
        self.btn_compress.config(state=tk.NORMAL)
        self.btn_compress_selected.config(state=tk.NORMAL)
        self.btn_cancel.config(state=tk.DISABLED, text="⏹️ 取消操作")
        
        if not was_canceled:
            self.progress_var.set(100)
            self.lbl_progress.config(text="全部处理完成!")
            msg = f"操作结束！成功: {succ} 个，失败: {fail} 个。"
            messagebox.showinfo("完成", msg)


def main():
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    app = GifCompressor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
