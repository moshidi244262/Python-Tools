# 依赖安装: pip install tkinterdnd2
# 注意: 运行本脚本需要系统已安装 FFmpeg 并配置好环境变量

import os
import sys
import json
import subprocess
import threading
import tkinter as tk
import re
from tkinter import ttk, messagebox, filedialog, scrolledtext
from datetime import datetime

# 尝试导入拖拽库，如果失败则弹出错误窗口而不是直接控制台闪退
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except ImportError:
    import tkinter as tk
    from tkinter import messagebox
    import sys
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("缺少依赖库", "错误: 未检测到 tkinterdnd2 库。\n\n请打开命令行(CMD)并运行以下命令安装:\npip install tkinterdnd2")
    sys.exit(1)

# 全局常量
SUPPORT_FORMATS = ('.mp4', '.avi', '.mkv', '.mov', '.flv', '.wmv', '.webm', '.m4v', '.mpeg', '.mpg')

class FFmpegController:
    """处理 FFmpeg 相关逻辑"""
    
    @staticmethod
    def check_ffmpeg():
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            subprocess.run(["ffmpeg", "-version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, startupinfo=startupinfo)
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            return False

    @staticmethod
    def format_duration(seconds):
        m, s = divmod(int(seconds), 60)
        h, m = divmod(m, 60)
        return f"{h:02d}:{m:02d}:{s:02d}" if h > 0 else f"{m:02d}:{s:02d}"

    @staticmethod
    def get_video_info(file_path):
        if not os.path.exists(file_path):
            return None
        
        cmd = ["ffprobe", "-v", "quiet", "-print_format", "json", "-show_format", "-show_streams", file_path]
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', errors='ignore', startupinfo=startupinfo)
            if result.returncode != 0: return None
            
            data = json.loads(result.stdout)
            video_stream = next((s for s in data.get('streams', []) if s['codec_type'] == 'video'), None)
            audio_stream = next((s for s in data.get('streams', []) if s['codec_type'] == 'audio'), None)
            
            if not video_stream: return None

            fps_str = video_stream.get('r_frame_rate', '0/1')
            num, den = map(int, fps_str.split('/'))
            fps = round(num / den, 2) if den != 0 else 0
            duration = float(data.get('format', {}).get('duration', 0))
            bitrate = int(data.get('format', {}).get('bit_rate', 0)) // 1000

            return {
                'path': file_path,
                'filename': os.path.basename(file_path),
                'width': int(video_stream.get('width', 0)),
                'height': int(video_stream.get('height', 0)),
                'fps': fps,
                'video_codec': video_stream.get('codec_name', 'N/A'),
                'duration': duration,
                'bitrate_kbps': bitrate,
                'audio_codec': audio_stream.get('codec_name', 'N/A') if audio_stream else 'N/A',
                'audio_bitrate': int(audio_stream.get('bit_rate', 0)) // 1000 if audio_stream else 0,
                'size': os.path.getsize(file_path)
            }
        except Exception as e:
            return None

    @staticmethod
    def build_command(input_path, output_path, params):
        cmd = ["ffmpeg", "-i", input_path]
        
        # 视频编码 (增加硬件加速选项)
        vcodec_map = {
            "H.264 (CPU)": "libx264",
            "H.265 (CPU)": "libx265",
            "H.264 (Nvidia GPU)": "h264_nvenc",
            "H.265 (Nvidia GPU)": "hevc_nvenc"
        }
        vcodec = vcodec_map.get(params['codec'], "libx264")
        cmd.extend(["-c:v", vcodec])
        
        # 码率控制
        if params['mode'] == "CRF" and "nvenc" not in vcodec:
            cmd.extend(["-crf", str(params['crf'])])
        elif params['mode'] == "CRF" and "nvenc" in vcodec:
            # NVENC 使用 cq 替代 crf
            cmd.extend(["-cq", str(params['crf']), "-b:v", "0"]) 
        else:
            cmd.extend(["-b:v", params['bitrate'], "-maxrate", params['bitrate'], "-bufsize", params['bitrate']])

        # 预设
        cmd.extend(["-preset", params['preset']])
        
        # 分辨率与帧率过滤
        filters = []
        if params['resolution_mode'] == "比例缩放":
            filters.append(f"scale=-2:{params['resolution_scale']}")
        elif params['resolution_mode'] == "自定义":
            filters.append(f"scale={params['resolution_custom'].replace('x', ':')}")
            
        if params['framerate_mode'] == "自定义":
            filters.append(f"fps={params['framerate_custom']}")
            
        if filters:
            cmd.extend(["-vf", ",".join(filters)])
            
        # 音频
        cmd.extend(["-c:a", "aac", "-b:a", params['audio_bitrate']])
        cmd.extend(["-y", output_path])
        return cmd

class VideoCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("视频批量压缩工具 (FFmpeg) - 优化版")
        self.root.geometry("1150x800")
        
        # 设置全局样式
        style = ttk.Style()
        style.theme_use('clam') # 使用更现代的默认主题
        style.configure('.', font=('Microsoft YaHei', 10))
        style.configure('TButton', padding=5)
        style.configure('TLabelframe.Label', font=('Microsoft YaHei', 10, 'bold'), foreground='#333333')
        
        self.file_list = []
        self.is_processing = False
        self.cancel_flag = False
        self.current_process = None
        
        if not FFmpegController.check_ffmpeg():
            messagebox.showerror("环境错误", "未检测到 FFmpeg！\n请下载并将其添加到系统环境变量中。")
            root.destroy()
            return

        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === 左侧：文件列表与进度 ===
        left_frame = ttk.LabelFrame(main_frame, text=" 📂 待处理文件 ", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btn_frame, text="➕ 添加文件", command=self.add_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="📁 添加文件夹", command=self.add_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🗑️ 删除选中", command=self.delete_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🧹 清空所有", command=self.clear_all).pack(side=tk.LEFT, padx=2)
        
        columns = ("filename", "resolution", "fps", "bitrate", "duration", "size", "status")
        self.tree = ttk.Treeview(left_frame, columns=columns, show="headings", selectmode="extended")
        for col, text, width in zip(columns, ["文件名", "分辨率", "帧率", "总码率", "时长", "大小", "状态"], [220, 90, 60, 80, 70, 80, 100]):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor="center" if col != "filename" else "w")
        
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self.on_drop)
        
        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # 进度区
        prog_frame = ttk.Frame(left_frame)
        prog_frame.pack(fill=tk.X, pady=(10, 0))
        self.progress_var = tk.StringVar(value="当前状态: 就绪")
        ttk.Label(prog_frame, textvariable=self.progress_var, font=('Microsoft YaHei', 9, 'bold')).pack(anchor="w")
        self.progress_bar = ttk.Progressbar(prog_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))

        # === 右侧：参数设置 ===
        right_frame = ttk.LabelFrame(main_frame, text=" ⚙️ 压缩参数 ", padding="15")
        right_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 布局辅助函数
        def add_row(parent, label_text, widget, row):
            ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky="e", pady=6, padx=(0, 5))
            widget.grid(row=row, column=1, sticky="ew", pady=6)

        # 1. 编码格式
        self.codec_var = tk.StringVar(value="H.264 (CPU)")
        codec_combo = ttk.Combobox(right_frame, textvariable=self.codec_var, values=["H.264 (CPU)", "H.265 (CPU)", "H.264 (Nvidia GPU)", "H.265 (Nvidia GPU)"], state="readonly", width=18)
        add_row(right_frame, "编码格式:", codec_combo, 0)

        # 2. 码率控制模式
        self.mode_var = tk.StringVar(value="CRF")
        mode_frame = ttk.Frame(right_frame)
        ttk.Radiobutton(mode_frame, text="固定质量(CRF)", variable=self.mode_var, value="CRF", command=self.toggle_mode).pack(side=tk.LEFT, padx=(0,10))
        ttk.Radiobutton(mode_frame, text="指定码率", variable=self.mode_var, value="Bitrate", command=self.toggle_mode).pack(side=tk.LEFT)
        add_row(right_frame, "控制模式:", mode_frame, 1)

        # 3. CRF / Bitrate 值
        self.crf_var = tk.StringVar(value="23")
        self.bitrate_var = tk.StringVar(value="2000k")
        self.param_entry = ttk.Entry(right_frame, textvariable=self.crf_var)
        add_row(right_frame, "质量/码率:", self.param_entry, 2)
        
        # 4. 分辨率
        self.res_mode_var = tk.StringVar(value="保持原始")
        self.res_scale_var = tk.StringVar(value="720")
        self.res_custom_var = tk.StringVar(value="1280x720")
        res_frame = ttk.Frame(right_frame)
        res_combo = ttk.Combobox(res_frame, textvariable=self.res_mode_var, values=["保持原始", "比例缩放", "自定义"], state="readonly", width=10)
        res_combo.pack(side=tk.LEFT, padx=(0,5))
        res_combo.bind("<<ComboboxSelected>>", lambda e: self.update_entry_states())
        self.res_scale_entry = ttk.Entry(res_frame, textvariable=self.res_scale_var, width=8)
        self.res_scale_entry.pack(side=tk.LEFT)
        add_row(right_frame, "分辨率:", res_frame, 3)
        
        # 5. 帧率
        self.fr_mode_var = tk.StringVar(value="保持原始")
        self.fr_val_var = tk.StringVar(value="30")
        fr_frame = ttk.Frame(right_frame)
        ttk.Radiobutton(fr_frame, text="原始", variable=self.fr_mode_var, value="保持原始", command=self.update_entry_states).pack(side=tk.LEFT, padx=(0,5))
        ttk.Radiobutton(fr_frame, text="自定义", variable=self.fr_mode_var, value="自定义", command=self.update_entry_states).pack(side=tk.LEFT)
        self.fr_entry = ttk.Entry(fr_frame, textvariable=self.fr_val_var, width=6)
        self.fr_entry.pack(side=tk.LEFT, padx=(5,0))
        add_row(right_frame, "帧率(FPS):", fr_frame, 4)

        # 6. 预设 & 音频
        self.preset_var = tk.StringVar(value="medium")
        add_row(right_frame, "编码预设:", ttk.Combobox(right_frame, textvariable=self.preset_var, values=["fast", "medium", "slow", "slower"], state="readonly"), 5)
        
        self.audio_bitrate_var = tk.StringVar(value="128k")
        add_row(right_frame, "音频码率:", ttk.Combobox(right_frame, textvariable=self.audio_bitrate_var, values=["64k", "96k", "128k", "192k", "256k"], state="readonly"), 6)

        # 7. 输出目录设置
        ttk.Separator(right_frame, orient="horizontal").grid(row=7, column=0, columnspan=2, sticky="ew", pady=10)
        self.out_dir_mode = tk.StringVar(value="同目录")
        ttk.Label(right_frame, text="输出位置:").grid(row=8, column=0, sticky="e", pady=2)
        out_dir_frame = ttk.Frame(right_frame)
        out_dir_frame.grid(row=8, column=1, sticky="ew")
        ttk.Radiobutton(out_dir_frame, text="源文件同目录", variable=self.out_dir_mode, value="同目录").pack(anchor="w")
        ttk.Radiobutton(out_dir_frame, text="指定文件夹", variable=self.out_dir_mode, value="自定义").pack(anchor="w", pady=(5,0))
        
        self.custom_dir_var = tk.StringVar()
        dir_sel_frame = ttk.Frame(right_frame)
        dir_sel_frame.grid(row=9, column=1, sticky="ew", pady=(2, 10))
        ttk.Entry(dir_sel_frame, textvariable=self.custom_dir_var, width=15).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(dir_sel_frame, text="浏览...", width=6, command=lambda: self.custom_dir_var.set(filedialog.askdirectory())).pack(side=tk.LEFT, padx=(2,0))

        # 操作按钮
        btn_action_frame = ttk.Frame(right_frame)
        btn_action_frame.grid(row=10, column=0, columnspan=2, sticky="ew", pady=10)
        btn_action_frame.columnconfigure(0, weight=1)
        btn_action_frame.columnconfigure(1, weight=1)
        
        self.start_btn = ttk.Button(btn_action_frame, text="▶ 开始压缩", command=self.start_compression)
        self.start_btn.grid(row=0, column=0, sticky="ew", padx=(0,5), ipady=5)
        self.stop_btn = ttk.Button(btn_action_frame, text="⏹ 停止", command=self.stop_compression, state="disabled")
        self.stop_btn.grid(row=0, column=1, sticky="ew", padx=(5,0), ipady=5)

        # === 底部：日志区 ===
        log_frame = ttk.LabelFrame(main_frame, text=" 📝 运行日志 ", padding="5")
        log_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(15, 0))
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=('Consolas', 9), state='disabled', bg='#f8f9fa')
        self.log_text.pack(fill=tk.X)

        self.update_entry_states()

    def toggle_mode(self):
        self.param_entry.config(textvariable=self.crf_var if self.mode_var.get() == "CRF" else self.bitrate_var)

    def update_entry_states(self):
        self.res_scale_entry.config(state='normal' if self.res_mode_var.get() != "保持原始" else 'disabled')
        if self.res_mode_var.get() == "自定义": self.res_scale_entry.config(textvariable=self.res_custom_var)
        elif self.res_mode_var.get() == "比例缩放": self.res_scale_entry.config(textvariable=self.res_scale_var)
        self.fr_entry.config(state='normal' if self.fr_mode_var.get() == "自定义" else 'disabled')

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def get_files_from_path(self, paths_str):
        paths = self.root.tk.splitlist(paths_str)
        valid_files = []
        for path in paths:
            if path.startswith('{') and path.endswith('}'): path = path[1:-1]
            if os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for f in files:
                        if f.lower().endswith(SUPPORT_FORMATS): valid_files.append(os.path.join(root, f))
            elif os.path.isfile(path) and path.lower().endswith(SUPPORT_FORMATS):
                valid_files.append(path)
        return valid_files

    def add_files(self):
        exts = " ".join([f"*{ext}" for ext in SUPPORT_FORMATS])
        files = filedialog.askopenfilename(filetypes=[("视频文件", exts), ("所有文件", "*.*")], multiple=True)
        self.add_to_list(files)

    def add_folder(self):
        folder = filedialog.askdirectory()
        if folder: self.add_to_list(self.get_files_from_path(folder))

    def on_drop(self, event):
        self.add_to_list(self.get_files_from_path(event.data))

    def add_to_list(self, files):
        new_count = 0
        for file_path in files:
            if any(f['path'] == file_path for f in self.file_list): continue
            info = FFmpegController.get_video_info(file_path)
            if info:
                self.file_list.append(info)
                self.tree.insert("", "end", iid=file_path, values=(
                    info['filename'], f"{info['width']}x{info['height']}", info['fps'],
                    f"{info['bitrate_kbps']}k", FFmpegController.format_duration(info['duration']),
                    f"{info['size'] / (1024 * 1024):.1f} MB", "待处理"
                ))
                new_count += 1
        if new_count > 0: self.log(f"成功导入 {new_count} 个视频文件。")

    def delete_selected(self):
        for item in self.tree.selection():
            self.tree.delete(item)
            self.file_list = [f for f in self.file_list if f['path'] != item]

    def clear_all(self):
        self.tree.delete(*self.tree.get_children())
        self.file_list.clear()

    def start_compression(self):
        if not self.file_list: return messagebox.showwarning("提示", "请先添加待处理的视频文件。")
        if self.out_dir_mode.get() == "自定义" and not self.custom_dir_var.get():
            return messagebox.showwarning("提示", "请选择自定义的输出文件夹。")
            
        self.is_processing = True
        self.cancel_flag = False
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        
        threading.Thread(target=self.process_files, daemon=True).start()

    def stop_compression(self):
        if messagebox.askyesno("确认", "确定要中断当前的压缩任务吗？"):
            self.cancel_flag = True
            self.log("正在停止压缩进程，请稍候...")
            if self.current_process:
                try: self.current_process.terminate()
                except: pass

    def get_output_path(self, original_path, filename):
        base, ext = os.path.splitext(filename)
        new_name = f"{base}_压缩版.mp4"
        
        if self.out_dir_mode.get() == "自定义":
            out_dir = self.custom_dir_var.get()
        else:
            out_dir = os.path.dirname(original_path)
            
        if not os.path.exists(out_dir): os.makedirs(out_dir, exist_ok=True)
        return os.path.join(out_dir, new_name)

    def process_files(self):
        total = len(self.file_list)
        self.root.after(0, self.log, f"========== 开始处理，共计 {total} 个文件 ==========")
        
        time_regex = re.compile(r"time=(\d{2}):(\d{2}):(\d{2}\.\d{2})")
        
        for idx, info in enumerate(self.file_list):
            if self.cancel_flag: break
            
            # 跳过已完成的
            if self.tree.item(info['path'], 'values')[6] == "已完成":
                continue

            output_path = self.get_output_path(info['path'], info['filename'])
            self.root.after(0, self.update_tree_status, info['path'], "▶ 压缩中")
            self.root.after(0, self.progress_var.set, f"正在处理 ({idx+1}/{total}): {info['filename']}")
            
            params = {
                'codec': self.codec_var.get(), 'mode': self.mode_var.get(),
                'crf': self.crf_var.get(), 'bitrate': self.bitrate_var.get(),
                'preset': self.preset_var.get(), 'resolution_mode': self.res_mode_var.get(),
                'resolution_scale': self.res_scale_var.get(), 'resolution_custom': self.res_custom_var.get(),
                'framerate_mode': self.fr_mode_var.get(), 'framerate_custom': self.fr_val_var.get(),
                'audio_bitrate': self.audio_bitrate_var.get()
            }
            
            cmd = FFmpegController.build_command(info['path'], output_path, params)
            total_duration = info['duration']
            
            try:
                startupinfo = subprocess.STARTUPINFO() if os.name == 'nt' else None
                if startupinfo: startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                
                # 使用 Popen 实时读取输出以获取进度
                self.current_process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='ignore', startupinfo=startupinfo)
                
                for line in self.current_process.stdout:
                    if self.cancel_flag: break
                    
                    match = time_regex.search(line)
                    if match and total_duration > 0:
                        h, m, s = float(match.group(1)), float(match.group(2)), float(match.group(3))
                        current_sec = h * 3600 + m * 60 + s
                        percent = min((current_sec / total_duration) * 100, 99.9)
                        self.root.after(0, self.update_progress, percent)
                
                self.current_process.wait()
                
                if self.cancel_flag:
                    self.root.after(0, self.update_tree_status, info['path'], "已取消")
                    break

                if self.current_process.returncode == 0:
                    new_size = os.path.getsize(output_path)
                    saved = (1 - new_size / info['size']) * 100 if info['size'] > 0 else 0
                    self.root.after(0, self.log, f"✅ 完成: {info['filename']} | 压缩后: {new_size/1024/1024:.1f}MB | 节省: {saved:.1f}%")
                    self.root.after(0, self.update_tree_status, info['path'], "已完成")
                    self.root.after(0, self.update_progress, 100.0)
                else:
                    self.root.after(0, self.log, f"❌ 失败: {info['filename']}")
                    self.root.after(0, self.update_tree_status, info['path'], "失败")
                    
            except Exception as e:
                self.root.after(0, self.log, f"⚠️ 错误: {info['filename']} - {str(e)}")
                self.root.after(0, self.update_tree_status, info['path'], "错误")
            finally:
                self.current_process = None

        self.root.after(0, self.finish_compression)

    def update_progress(self, percent):
        self.progress_bar['value'] = percent

    def update_tree_status(self, item_id, status):
        try:
            vals = list(self.tree.item(item_id, 'values'))
            vals[6] = status # 状态现在在第7列(索引6)
            self.tree.item(item_id, values=vals)
        except: pass

    def finish_compression(self):
        self.is_processing = False
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        
        if self.cancel_flag:
            self.progress_var.set("状态: 任务已手动终止")
            self.log("========== 任务被手动终止 ==========")
        else:
            self.progress_var.set("状态: 所有任务处理完成")
            self.log("========== 队列处理完毕 ==========")
            messagebox.showinfo("完成", "队列中的视频压缩任务已处理完毕！")

if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1) # 解决Windows下界面模糊问题
    except: pass

    root = TkinterDnD.Tk()
    app = VideoCompressorApp(root)
    root.mainloop()
