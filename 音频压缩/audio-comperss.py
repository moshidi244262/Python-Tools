# 依赖安装: pip install mutagen tkinterdnd2
# 优化说明：修复UI线程安全、加入多线程并发、新增进度条、自定义输出路径、UI全面美化、恢复详细参数显示

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import traceback
from pathlib import Path
import subprocess
import shutil
import queue
from concurrent.futures import ThreadPoolExecutor
import time

# ==========================================
# 安全导入第三方库
# ==========================================
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    import mutagen
except ImportError as e:
    root = tk.Tk()
    root.withdraw()
    error_msg = f"启动失败，缺少第三方库：\n{str(e)}\n\n请运行：pip install mutagen tkinterdnd2"
    messagebox.showerror("环境配置错误", error_msg)
    sys.exit(1)


class AudioCompressorApp:
    SUPPORTED_FORMATS = ('.flac', '.wav', '.mp3', '.m4a', '.aac', '.ogg')
    
    def __init__(self, root):
        self.root = root
        self.root.title("✨ 极速音频压缩工具 v2.1 (多线程+详细参数+FFmpeg引擎)")
        self.root.geometry("1150x750")
        self.root.minsize(1000, 650)
        
        # 启用现代主题
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')
            
        # 线程安全队列
        self.ui_queue = queue.Queue()
        
        # 数据存储
        self.audio_files = {}  # {filepath: info_dict}
        
        # 变量
        self.bitrate_var = tk.StringVar(value="128")
        self.sample_rate_var = tk.StringVar(value="44100")
        self.channels_var = tk.StringVar(value="2")
        self.output_dir_var = tk.StringVar(value=str(Path.home() / "Desktop" / "压缩音频输出"))
        
        # 标志位
        self.is_compressing = False
        self.cancel_flag = False
        self.completed_count = 0
        self.total_count = 0
        
        self._setup_styles()
        self._setup_ui()
        
        # 启动UI事件循环处理队列
        self.root.after(100, self._process_ui_queue)
        
    def _setup_styles(self):
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Microsoft YaHei', 10, 'bold'), foreground='#2c3e50')
        style.configure('TButton', font=('Microsoft YaHei', 9), padding=5)
        style.configure('Action.TButton', font=('Microsoft YaHei', 9, 'bold'))
        style.configure('Treeview', rowheight=25, font=('Microsoft YaHei', 9))
        style.configure('Treeview.Heading', font=('Microsoft YaHei', 9, 'bold'), background='#e1e8ed')
        style.map('Treeview', background=[('selected', '#3498db')])
        
    def _setup_ui(self):
        # ========== 顶部参数与路径设置 ==========
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill='x', padx=15, pady=10)
        
        # 参数设置区
        param_frame = ttk.LabelFrame(top_frame, text=" ⚙️ 压缩参数 ", padding=10)
        param_frame.pack(side='left', fill='y', expand=False, padx=(0, 10))
        
        ttk.Label(param_frame, text="比特率:").grid(row=0, column=0, padx=5, pady=5)
        ttk.Combobox(param_frame, textvariable=self.bitrate_var, 
                    values=["64", "96", "128", "160", "192", "256", "320"], 
                    width=6, state='readonly').grid(row=0, column=1, pady=5)
        ttk.Label(param_frame, text="kbps").grid(row=0, column=2, padx=(0, 15))
        
        ttk.Label(param_frame, text="采样率:").grid(row=0, column=3, padx=5, pady=5)
        ttk.Combobox(param_frame, textvariable=self.sample_rate_var,
                    values=["22050", "32000", "44100", "48000"], 
                    width=8, state='readonly').grid(row=0, column=4, pady=5)
        ttk.Label(param_frame, text="Hz").grid(row=0, column=5, padx=(0, 15))
        
        ttk.Label(param_frame, text="声道:").grid(row=0, column=6, padx=5, pady=5)
        ttk.Combobox(param_frame, textvariable=self.channels_var,
                    values=["1 (单声道)", "2 (立体声)"], 
                    width=12, state='readonly').grid(row=0, column=7, pady=5)

        # 路径设置区
        path_frame = ttk.LabelFrame(top_frame, text=" 📂 输出目录 ", padding=10)
        path_frame.pack(side='left', fill='both', expand=True)
        
        ttk.Entry(path_frame, textvariable=self.output_dir_var).pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(path_frame, text="浏览...", command=self._choose_output_dir, width=8).pack(side='left', padx=2)
        ttk.Button(path_frame, text="打开目录", command=self._open_output_dir, width=10).pack(side='left', padx=2)

        # ========== 按钮操作区 ==========
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill='x', padx=15, pady=5)
        
        ttk.Button(btn_frame, text="➕ 添加文件", command=self._select_files).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="📁 添加文件夹", command=self._select_folder).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="🗑 清空列表", command=self._clear_list).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="❌ 移除选中", command=self._delete_selected).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="📝 清空日志", command=self._clear_log).pack(side='left', padx=3)
        
        self.stop_btn = ttk.Button(btn_frame, text="⏹ 停止", command=self._stop_compress, state='disabled')
        self.stop_btn.pack(side='right', padx=3)
        
        self.compress_btn = ttk.Button(btn_frame, text="🚀 开始批量压缩", style='Action.TButton', command=self._start_compress)
        self.compress_btn.pack(side='right', padx=10)

        # ========== 列表区域 ==========
        list_frame = ttk.Frame(self.root)
        list_frame.pack(fill='both', expand=True, padx=15, pady=5)
        
        # 增加参数显示列
        columns = ('filename', 'orig_format', 'orig_bitrate', 'orig_sample_rate', 'orig_channels', 'orig_size', 'status')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', selectmode='extended', height=12)
        
        self.tree.heading('filename', text='文件名', anchor='w')
        self.tree.heading('orig_format', text='格式', anchor='center')
        self.tree.heading('orig_bitrate', text='原比特率', anchor='center')
        self.tree.heading('orig_sample_rate', text='原采样率', anchor='center')
        self.tree.heading('orig_channels', text='原声道', anchor='center')
        self.tree.heading('orig_size', text='文件大小', anchor='center')
        self.tree.heading('status', text='处理状态', anchor='center')
        
        self.tree.column('filename', width=300, anchor='w')
        self.tree.column('orig_format', width=60, anchor='center')
        self.tree.column('orig_bitrate', width=80, anchor='center')
        self.tree.column('orig_sample_rate', width=80, anchor='center')
        self.tree.column('orig_channels', width=60, anchor='center')
        self.tree.column('orig_size', width=80, anchor='center')
        self.tree.column('status', width=220, anchor='center')
        
        self.tree.tag_configure('evenrow', background='#f9f9f9')
        self.tree.tag_configure('oddrow', background='#ffffff')
        self.tree.tag_configure('success', foreground='#27ae60', font=('Microsoft YaHei', 9, 'bold'))
        self.tree.tag_configure('error', foreground='#e74c3c')
        self.tree.tag_configure('processing', foreground='#f39c12')
        
        scroll_y = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        scroll_x = ttk.Scrollbar(list_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree.pack(side='top', fill='both', expand=True)
        scroll_y.pack(side='right', fill='y')
        
        # 拖拽绑定
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self._on_drop)
        
        # ========== 进度与日志区 ==========
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill='x', padx=15, pady=10)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(bottom_frame, orient='horizontal', mode='determinate', variable=self.progress_var)
        self.progress_bar.pack(fill='x', pady=(0, 5))
        
        # 状态文本
        self.status_label = ttk.Label(bottom_frame, text="就绪 | 等待导入文件...", style='Title.TLabel')
        self.status_label.pack(anchor='w')

        # 日志框
        log_frame = ttk.LabelFrame(bottom_frame, text=" 运行日志 ", padding=5)
        log_frame.pack(fill='both', expand=True, pady=5)
        self.log_text = tk.Text(log_frame, height=8, font=('Consolas', 9), bg='#f8f9fa')
        log_scroll = ttk.Scrollbar(log_frame, orient='vertical', command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scroll.pack(side='right', fill='y')
        
        self.log_text.tag_configure('info', foreground='#333333')
        self.log_text.tag_configure('success', foreground='#27ae60')
        self.log_text.tag_configure('error', foreground='#e74c3c')
        self.log_text.tag_configure('warning', foreground='#d35400')

    # ================= UI 线程安全更新机制 =================
    def _process_ui_queue(self):
        """处理来自子线程的 UI 更新请求"""
        try:
            while True:
                task = self.ui_queue.get_nowait()
                action = task[0]
                
                if action == 'log':
                    msg, level = task[1], task[2]
                    self.log_text.insert('end', msg + '\n', level)
                    self.log_text.see('end')
                elif action == 'update_tree':
                    item_id, values, tags = task[1], task[2], task[3]
                    if self.tree.exists(item_id):
                        self.tree.item(item_id, values=values, tags=tags)
                elif action == 'update_progress':
                    self.completed_count += 1
                    percent = (self.completed_count / self.total_count) * 100
                    self.progress_var.set(percent)
                    self.status_label.config(text=f"进度: {self.completed_count}/{self.total_count} ({percent:.1f}%)")
                elif action == 'finish':
                    self._on_compress_finish(task[1], task[2])
                    
                self.ui_queue.task_done()
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._process_ui_queue)

    def _log_safe(self, message, level='info'):
        """线程安全的日志记录"""
        self.ui_queue.put(('log', message, level))
        
    def _update_tree_safe(self, item_id, values, tags=()):
        """线程安全的 Treeview 更新"""
        self.ui_queue.put(('update_tree', item_id, values, tags))

    # ================= 业务逻辑 =================
    def _format_size(self, size_bytes):
        try:
            size_bytes = float(size_bytes)
            for unit in ['B', 'KB', 'MB', 'GB']:
                if size_bytes < 1024:
                    return f"{size_bytes:.1f} {unit}"
                size_bytes /= 1024
            return f"{size_bytes:.1f} TB"
        except:
            return "N/A"

    def _get_audio_info(self, filepath):
        """使用 mutagen 提取音频参数信息"""
        info = {
            'bitrate': '未知',
            'sample_rate': '未知',
            'channels': '未知'
        }
        try:
            audio = mutagen.File(filepath, easy=True)
            if audio is not None:
                audio_info = audio.info
                if hasattr(audio_info, 'bitrate') and audio_info.bitrate:
                    info['bitrate'] = f"{int(audio_info.bitrate / 1000)} kbps"
                if hasattr(audio_info, 'sample_rate') and audio_info.sample_rate:
                    info['sample_rate'] = f"{int(audio_info.sample_rate)} Hz"
                if hasattr(audio_info, 'channels'):
                    info['channels'] = str(audio_info.channels)
        except Exception:
            pass
        return info

    def _build_tree_values(self, info, status):
        """生成符合 Treeview 列格式的元组"""
        return (
            info['filename'], 
            info['format'], 
            info.get('bitrate', '未知'),
            info.get('sample_rate', '未知'),
            info.get('channels', '未知'),
            info['size'], 
            status
        )

    def _choose_output_dir(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir_var.set(folder)
            
    def _open_output_dir(self):
        path = self.output_dir_var.get()
        if os.path.exists(path):
            os.startfile(path) if sys.platform == 'win32' else subprocess.run(['open', path])
        else:
            messagebox.showinfo("提示", "输出目录尚未创建！")

    def _add_file(self, filepath):
        filepath = os.path.abspath(filepath)
        if filepath in self.audio_files or not os.path.isfile(filepath):
            return False
            
        ext = os.path.splitext(filepath)[1].lower()
        if ext not in self.SUPPORTED_FORMATS:
            return False
            
        size = os.path.getsize(filepath)
        audio_params = self._get_audio_info(filepath)
        
        info = {
            'filepath': filepath,
            'filename': os.path.basename(filepath),
            'format': ext[1:].upper(),
            'size': self._format_size(size),
            'size_bytes': size,
            'bitrate': audio_params['bitrate'],
            'sample_rate': audio_params['sample_rate'],
            'channels': audio_params['channels']
        }
        
        self.audio_files[filepath] = info
        self._refresh_tree()
        return True

    def _refresh_tree(self):
        """重新渲染列表（处理斑马线及全列数据）"""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for i, (filepath, info) in enumerate(self.audio_files.items()):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            values = self._build_tree_values(info, '等待压缩')
            item_id = self.tree.insert('', 'end', values=values, tags=(tag,))
            self.audio_files[filepath]['tree_id'] = item_id
            
        self.status_label.config(text=f"已加载 {len(self.audio_files)} 个文件")

    def _add_files(self, filepaths):
        count = sum(1 for p in filepaths if self._add_file(p))
        if count > 0: self._log_safe(f"✅ 成功添加 {count} 个文件", 'info')

    def _select_files(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("音频文件", "*.flac *.wav *.mp3 *.m4a *.aac *.ogg"), ("所有文件", "*.*")])
        if filepaths: self._add_files(filepaths)

    def _select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            count = 0
            for root, _, files in os.walk(folder):
                for f in files:
                    if self._add_file(os.path.join(root, f)): count += 1
            if count > 0: self._log_safe(f"📂 从文件夹导入了 {count} 个文件", 'info')

    def _on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        file_count = 0
        for path in paths:
            if os.path.isfile(path):
                if self._add_file(path): file_count += 1
            elif os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for f in files:
                        if self._add_file(os.path.join(root, f)): file_count += 1
        if file_count > 0: self._log_safe(f"📥 拖拽导入了 {file_count} 个文件", 'info')

    def _clear_list(self):
        self.audio_files.clear()
        self._refresh_tree()
        self._log_safe("🗑 已清空列表")

    def _clear_log(self):
        """清空日志框内容"""
        self.log_text.delete(1.0, 'end')

    def _delete_selected(self):
        selected = self.tree.selection()
        for item_id in selected:
            for filepath, info in list(self.audio_files.items()):
                if info.get('tree_id') == item_id:
                    del self.audio_files[filepath]
                    break
        self._refresh_tree()

    def _stop_compress(self):
        if self.is_compressing:
            self.cancel_flag = True
            self.status_label.config(text="正在中止任务，请稍候...")
            self.stop_btn.config(state='disabled')

    def _start_compress(self):
        if not self.audio_files:
            messagebox.showwarning("提示", "请先添加文件！")
            return
            
        out_dir = self.output_dir_var.get().strip()
        if not out_dir:
            messagebox.showwarning("提示", "请设置输出目录！")
            return
            
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        
        self.is_compressing = True
        self.cancel_flag = False
        self.compress_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        
        self.total_count = len(self.audio_files)
        self.completed_count = 0
        self.progress_var.set(0)
        
        self._clear_log()
        self._log_safe(f"🚀 开始执行并发压缩任务，共 {self.total_count} 个文件...", 'info')
        self._log_safe(f"目标参数: {self.bitrate_var.get()}kbps | {self.sample_rate_var.get()}Hz | 输出: {out_dir}\n", 'info')
        
        # 启动后台主控线程
        threading.Thread(target=self._run_compression_pool, daemon=True).start()

    def _run_compression_pool(self):
        """线程池管理器"""
        success_count = 0
        fail_count = 0
        
        # 根据CPU核心数动态分配线程，最多开 6 个并发防止卡死机器
        max_workers = min(os.cpu_count() or 4, 6) 
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = []
            for filepath, info in self.audio_files.items():
                if self.cancel_flag: break
                futures.append(executor.submit(self._compress_single_task, filepath, info))
                
            for future in futures:
                if self.cancel_flag: break
                try:
                    result = future.result()
                    if result: success_count += 1
                    else: fail_count += 1
                except Exception as e:
                    fail_count += 1
                    self._log_safe(f"未知异常: {str(e)}", 'error')
                
                self.ui_queue.put(('update_progress',))

        self.ui_queue.put(('finish', success_count, fail_count))

    def _compress_single_task(self, filepath, info):
        """在工作线程中执行 FFmpeg 调用"""
        if self.cancel_flag:
            return False
            
        tree_id = info.get('tree_id')
        filename = info['filename']
        self._update_tree_safe(tree_id, self._build_tree_values(info, "🔄 处理中..."), ('processing',))
        
        out_dir = Path(self.output_dir_var.get())
        base_name = os.path.splitext(filename)[0]
        out_path = out_dir / f"{base_name}.mp3"
        
        # 处理重名
        counter = 1
        while out_path.exists():
            out_path = out_dir / f"{base_name}_{counter}.mp3"
            counter += 1

        cmd = [
            "ffmpeg", "-y", "-v", "error",
            "-i", str(filepath),
            "-ar", self.sample_rate_var.get(),
            "-ac", str(int(self.channels_var.get()[0])),
            "-b:a", f"{self.bitrate_var.get()}k",
            "-map_metadata", "0", 
            "-id3v2_version", "3", # 修复Windows下MP3封面和歌手不显示的问题
            str(out_path)
        ]

        creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        
        try:
            start_time = time.time()
            process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=creationflags)
            
            if process.returncode != 0:
                self._update_tree_safe(tree_id, self._build_tree_values(info, "❌ 失败"), ('error',))
                self._log_safe(f"[{filename}] FFmpeg 报错: {process.stderr.strip()}", 'error')
                return False
                
            cost_time = time.time() - start_time
            new_size = os.path.getsize(out_path)
            orig_size = info['size_bytes']
            ratio = (1 - new_size/orig_size)*100 if orig_size > 0 else 0
            
            status_text = f"✅ 完成 ({self._format_size(new_size)} | 压缩率: {ratio:.1f}% | 耗时:{cost_time:.1f}s)"
            # 更新列表格式：格式改为 MP3，大小保持原样以作对比，状态更新为成功
            finished_info = info.copy()
            finished_info['format'] = "MP3"
            self._update_tree_safe(tree_id, self._build_tree_values(finished_info, status_text), ('success',))
            self._log_safe(f"✓ 成功: {filename} -> {status_text}", 'success')
            return True
            
        except Exception as e:
            self._update_tree_safe(tree_id, self._build_tree_values(info, "❌ 异常"), ('error',))
            self._log_safe(f"[{filename}] 程序异常: {str(e)}", 'error')
            return False

    def _on_compress_finish(self, success, fail):
        """恢复UI状态"""
        self.is_compressing = False
        self.compress_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        
        if self.cancel_flag:
            self.status_label.config(text="任务已用户中止")
            messagebox.showinfo("中止", "压缩任务已被手动中止。")
        else:
            self.status_label.config(text=f"完成！成功: {success} 个，失败: {fail} 个")
            messagebox.showinfo("处理完成", f"批量处理结束！\n成功: {success}\n失败: {fail}")


if __name__ == '__main__':
    root = TkinterDnD.Tk()
    
    if shutil.which("ffmpeg") is None:
        messagebox.showerror("缺少组件", "未检测到 FFmpeg！\n请将 ffmpeg.exe 放入当前文件夹或加入环境变量。")
        sys.exit(1)
        
    app = AudioCompressorApp(root)
    root.update_idletasks()
    x = (root.winfo_screenwidth() - root.winfo_width()) // 2
    y = (root.winfo_screenheight() - root.winfo_height()) // 2
    root.geometry(f'+{x}+{y}')
    root.mainloop()
