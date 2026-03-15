# 依赖安装: pip install tkinterdnd2
# 注意: 运行本脚本需要系统已安装 FFmpeg 并配置好环境变量

import os
import sys
import json
import subprocess
import threading
import tkinter as tk
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
    root.withdraw() # 隐藏主窗口
    messagebox.showerror("缺少依赖库", "错误: 未检测到 tkinterdnd2 库。\n\n请打开命令行(CMD)并运行以下命令安装:\npip install tkinterdnd2")
    sys.exit(1)

# 全局常量
SUPPORT_FORMATS = ('.mp4', '.avi', '.mkv', '.mov', '.flv', '.wmv', '.webm', '.m4v', '.mpeg', '.mpg')
OUTPUT_DIR_NAME = "压缩视频"

class FFmpegController:
    """处理 FFmpeg 相关逻辑"""
    
    @staticmethod
    def check_ffmpeg():
        try:
            # 隐藏 Windows 下的黑框
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
        """将秒数格式化为 HH:MM:SS 或 MM:SS"""
        m, s = divmod(int(seconds), 60)
        h, m = divmod(m, 60)
        return f"{h:02d}:{m:02d}:{s:02d}" if h > 0 else f"{m:02d}:{s:02d}"

    @staticmethod
    def get_video_info(file_path):
        """使用 ffprobe 获取视频详细信息"""
        if not os.path.exists(file_path):
            return None
        
        cmd = [
            "ffprobe",
            "-v", "quiet",
            "-print_format", "json",
            "-show_format",
            "-show_streams",
            file_path
        ]
        
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', errors='ignore', startupinfo=startupinfo)
            if result.returncode != 0:
                return None
            
            data = json.loads(result.stdout)
            video_stream = None
            audio_stream = None
            
            for stream in data.get('streams', []):
                if stream['codec_type'] == 'video' and video_stream is None:
                    video_stream = stream
                elif stream['codec_type'] == 'audio' and audio_stream is None:
                    audio_stream = stream
            
            if not video_stream:
                return None

            # 解析帧率 (可能是 "30/1" 或 "30000/1001")
            fps_str = video_stream.get('r_frame_rate', '0/1')
            num, den = map(int, fps_str.split('/'))
            fps = round(num / den, 2) if den != 0 else 0

            # 解析时长
            duration = float(data.get('format', {}).get('duration', 0))

            # 解析码率
            bitrate = int(data.get('format', {}).get('bit_rate', 0)) // 1000

            info = {
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
            return info
        except Exception as e:
            print(f"Error parsing {file_path}: {e}")
            return None

    @staticmethod
    def build_command(input_path, output_path, params):
        """构建 ffmpeg 压缩命令"""
        cmd = ["ffmpeg", "-i", input_path]
        
        # 视频编码
        vcodec = "libx264" if params['codec'] == "H.264" else "libx265"
        cmd.extend(["-c:v", vcodec])
        
        # 码率控制
        if params['mode'] == "CRF":
            cmd.extend(["-crf", str(params['crf'])])
        else:
            cmd.extend(["-b:v", params['bitrate']])
            cmd.extend(["-maxrate", params['bitrate']]) # 稍微限制峰值

        # 预设
        cmd.extend(["-preset", params['preset']])
        
        # 分辨率
        scale_filter = None
        res_mode = params['resolution_mode']
        if res_mode == "比例缩放":
            target_h = int(params['resolution_scale'])
            scale_filter = f"scale=-2:{target_h}"
        elif res_mode == "自定义":
            w, h = params['resolution_custom'].split('x')
            scale_filter = f"scale={w}:{h}"
        
        # 帧率
        fr_mode = params['framerate_mode']
        if fr_mode != "保持原始":
            fps_val = params['framerate_custom']
            if scale_filter:
                scale_filter += f",fps={fps_val}"
            else:
                scale_filter = f"fps={fps_val}"
        
        if scale_filter:
            cmd.extend(["-vf", scale_filter])
            
        # 音频参数
        cmd.extend(["-c:a", "aac", "-b:a", params['audio_bitrate']])
        
        # 覆盖输出
        cmd.extend(["-y", output_path])
        return cmd

class VideoCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("视频批量压缩工具 (FFmpeg)")
        self.root.geometry("1100x750")
        
        self.file_list = [] # 存储文件信息字典
        self.is_processing = False
        
        # 检查 FFmpeg
        if not FFmpegController.check_ffmpeg():
            messagebox.showerror("错误", "未检测到 FFmpeg 或 ffprobe！\n请下载安装 FFmpeg 并将其添加到系统环境变量中。")
            root.destroy()
            return

        self.setup_ui()
        
    def setup_ui(self):
        # 主布局
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧：文件列表
        left_frame = ttk.LabelFrame(main_frame, text="待处理文件", padding="5")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 列表工具栏
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(btn_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="添加文件夹", command=self.add_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="删除选中", command=self.delete_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="清空所有", command=self.clear_all).pack(side=tk.LEFT, padx=2)
        
        # 文件列表 Treeview (根据要求四增加列)
        columns = ("filename", "resolution", "fps", "bitrate", "vcodec", "abitrate", "duration", "size", "status")
        self.tree = ttk.Treeview(left_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("filename", text="文件名")
        self.tree.heading("resolution", text="分辨率")
        self.tree.heading("fps", text="帧率")
        self.tree.heading("bitrate", text="视频码率")
        self.tree.heading("vcodec", text="视频编码")
        self.tree.heading("abitrate", text="音频码率")
        self.tree.heading("duration", text="时长")
        self.tree.heading("size", text="大小")
        self.tree.heading("status", text="状态")
        
        self.tree.column("filename", width=180, anchor="w")
        self.tree.column("resolution", width=90, anchor="center")
        self.tree.column("fps", width=60, anchor="center")
        self.tree.column("bitrate", width=80, anchor="center")
        self.tree.column("vcodec", width=70, anchor="center")
        self.tree.column("abitrate", width=70, anchor="center")
        self.tree.column("duration", width=70, anchor="center")
        self.tree.column("size", width=80, anchor="center")
        self.tree.column("status", width=80, anchor="center")
        
        # 拖拽支持
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self.on_drop)
        
        # 进度条
        self.progress_var = tk.StringVar(value="就绪")
        self.progress_bar = ttk.Progressbar(left_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=5)
        ttk.Label(left_frame, textvariable=self.progress_var).pack(fill=tk.X)

        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # 右侧：参数设置
        right_frame = ttk.LabelFrame(main_frame, text="压缩参数", padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        
        # 1. 编码格式
        ttk.Label(right_frame, text="编码格式:").grid(row=0, column=0, sticky="w", pady=2)
        self.codec_var = tk.StringVar(value="H.264")
        codec_combo = ttk.Combobox(right_frame, textvariable=self.codec_var, values=["H.264", "H.265"], state="readonly", width=15)
        codec_combo.grid(row=0, column=1, sticky="ew", pady=2)

        # 2. 码率控制模式
        ttk.Label(right_frame, text="控制模式:").grid(row=1, column=0, sticky="w", pady=2)
        self.mode_var = tk.StringVar(value="CRF")
        mode_frame = ttk.Frame(right_frame)
        mode_frame.grid(row=1, column=1, sticky="ew")
        ttk.Radiobutton(mode_frame, text="CRF", variable=self.mode_var, value="CRF", command=self.toggle_mode).pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="Bitrate", variable=self.mode_var, value="Bitrate", command=self.toggle_mode).pack(side=tk.LEFT)

        # 3. CRF / Bitrate 值
        ttk.Label(right_frame, text="质量参数:").grid(row=2, column=0, sticky="w", pady=2)
        self.crf_var = tk.StringVar(value="23")
        self.bitrate_var = tk.StringVar(value="2000k")
        
        self.param_entry = ttk.Entry(right_frame, textvariable=self.crf_var, width=15)
        self.param_entry.grid(row=2, column=1, sticky="ew", pady=2)
        
        # 4. 分辨率
        ttk.Label(right_frame, text="分辨率:").grid(row=3, column=0, sticky="w", pady=2)
        self.res_mode_var = tk.StringVar(value="保持原始")
        self.res_scale_var = tk.StringVar(value="720")
        self.res_custom_var = tk.StringVar(value="1280x720")
        
        res_frame = ttk.Frame(right_frame)
        res_frame.grid(row=3, column=1, sticky="ew")
        res_combo = ttk.Combobox(res_frame, textvariable=self.res_mode_var, values=["保持原始", "比例缩放", "自定义"], state="readonly", width=15)
        res_combo.pack(fill=tk.X)
        res_combo.bind("<<ComboboxSelected>>", lambda e: self.update_entry_states())

        self.res_scale_entry = ttk.Entry(res_frame, textvariable=self.res_scale_var, width=5)
        self.res_scale_entry.pack(side=tk.LEFT, padx=2)
        ttk.Label(res_frame, text="p").pack(side=tk.LEFT)
        
        # 5. 帧率
        ttk.Label(right_frame, text="帧率:").grid(row=4, column=0, sticky="w", pady=2)
        self.fr_mode_var = tk.StringVar(value="保持原始")
        self.fr_val_var = tk.StringVar(value="30")
        fr_frame = ttk.Frame(right_frame)
        fr_frame.grid(row=4, column=1, sticky="ew")
        ttk.Radiobutton(fr_frame, text="原始", variable=self.fr_mode_var, value="保持原始", command=self.update_entry_states).pack(side=tk.LEFT)
        ttk.Radiobutton(fr_frame, text="自定义", variable=self.fr_mode_var, value="自定义", command=self.update_entry_states).pack(side=tk.LEFT)
        self.fr_entry = ttk.Entry(fr_frame, textvariable=self.fr_val_var, width=5)
        self.fr_entry.pack(side=tk.LEFT, padx=2)

        # 6. 预设
        ttk.Label(right_frame, text="编码预设:").grid(row=5, column=0, sticky="w", pady=2)
        self.preset_var = tk.StringVar(value="medium")
        preset_combo = ttk.Combobox(right_frame, textvariable=self.preset_var, 
                                    values=["ultrafast", "superfast", "veryfast", "faster", "fast", "medium", "slow", "slower", "veryslow"], 
                                    state="readonly", width=15)
        preset_combo.grid(row=5, column=1, sticky="ew", pady=2)

        # 7. 音频
        ttk.Label(right_frame, text="音频码率:").grid(row=6, column=0, sticky="w", pady=2)
        self.audio_bitrate_var = tk.StringVar(value="128k")
        audio_combo = ttk.Combobox(right_frame, textvariable=self.audio_bitrate_var, values=["64k", "128k", "192k", "256k", "320k"], state="readonly", width=15)
        audio_combo.grid(row=6, column=1, sticky="ew", pady=2)
        
        # 操作按钮
        ttk.Separator(right_frame, orient="horizontal").grid(row=7, column=0, columnspan=2, sticky="ew", pady=10)
        self.start_btn = ttk.Button(right_frame, text="开始压缩", command=self.start_compression)
        self.start_btn.grid(row=8, column=0, columnspan=2, sticky="ew", pady=5)
        
        ttk.Button(right_frame, text="清空日志", command=self.clear_log).grid(row=9, column=0, columnspan=2, sticky="ew", pady=5)

        # 底部：日志区
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding="5")
        log_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, state='disabled')
        self.log_text.pack(fill=tk.X)

        self.update_entry_states()

    def toggle_mode(self):
        if self.mode_var.get() == "CRF":
            self.param_entry.config(textvariable=self.crf_var)
        else:
            self.param_entry.config(textvariable=self.bitrate_var)

    def update_entry_states(self):
        # 分辨率控件状态
        state_scale = 'normal' if self.res_mode_var.get() == "比例缩放" else 'disabled'
        self.res_scale_entry.config(state=state_scale)
        
        # 帧率控件状态
        state_fr = 'normal' if self.fr_mode_var.get() == "自定义" else 'disabled'
        self.fr_entry.config(state=state_fr)

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def get_files_from_path(self, paths_str):
        """解析拖拽或选择的路径"""
        paths = self.root.tk.splitlist(paths_str)
        valid_files = []
        
        for path in paths:
            # 处理 Windows 路径可能带有的花括号
            if path.startswith('{') and path.endswith('}'):
                path = path[1:-1]
            
            if os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for f in files:
                        if f.lower().endswith(SUPPORT_FORMATS):
                            valid_files.append(os.path.join(root, f))
            elif os.path.isfile(path):
                if path.lower().endswith(SUPPORT_FORMATS):
                    valid_files.append(path)
        return valid_files

    def add_files(self):
        # 修复了无法选择文件的严重Bug：将支持的格式拼为标准Tkinter可用通配符
        exts = " ".join([f"*{ext}" for ext in SUPPORT_FORMATS])
        files = filedialog.askopenfilename(
            filetypes=[("视频文件", exts), ("所有文件", "*.*")], 
            multiple=True
        )
        self.add_to_list(files)

    def add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            files = []
            for root, _, fs in os.walk(folder):
                for f in fs:
                    if f.lower().endswith(SUPPORT_FORMATS):
                        files.append(os.path.join(root, f))
            self.add_to_list(files)

    def on_drop(self, event):
        files = self.get_files_from_path(event.data)
        self.add_to_list(files)

    def add_to_list(self, files):
        new_count = 0
        for file_path in files:
            # 去重
            if any(f['path'] == file_path for f in self.file_list):
                continue
            
            info = FFmpegController.get_video_info(file_path)
            if info:
                self.file_list.append(info)
                size_mb = info['size'] / (1024 * 1024)
                duration_str = FFmpegController.format_duration(info['duration'])
                audio_br_str = f"{info['audio_bitrate']}k" if info['audio_bitrate'] > 0 else "N/A"
                
                self.tree.insert("", "end", iid=file_path, values=(
                    info['filename'],
                    f"{info['width']}x{info['height']}",
                    f"{info['fps']}",
                    f"{info['bitrate_kbps']}k",
                    info['video_codec'].upper(),
                    audio_br_str,
                    duration_str,
                    f"{size_mb:.2f} MB",
                    "待处理"
                ))
                new_count += 1
            else:
                self.log(f"跳过无效文件: {os.path.basename(file_path)}")
        
        if new_count > 0:
            self.log(f"已添加 {new_count} 个文件。")

    def delete_selected(self):
        selected = self.tree.selection()
        for item in selected:
            self.tree.delete(item)
            self.file_list = [f for f in self.file_list if f['path'] == item]

    def clear_all(self):
        self.tree.delete(*self.tree.get_children())
        self.file_list = []
        self.log("列表已清空。")

    def check_parameters(self):
        """检查参数智能提示 (补充了帧率校验)"""
        issues = []
        
        # 获取目标参数
        target_h = None
        if self.res_mode_var.get() == "比例缩放":
            target_h = int(self.res_scale_var.get())
        elif self.res_mode_var.get() == "自定义":
            parts = self.res_custom_var.get().split('x')
            if len(parts) == 2:
                target_h = int(parts[1])
        
        target_bitrate = None
        if self.mode_var.get() == "Bitrate":
            br_str = self.bitrate_var.get().lower().replace('k', '')
            try: target_bitrate = int(br_str) * 1000 # 转为 bits (approx)
            except ValueError: pass
            
        target_fps = None
        if self.fr_mode_var.get() == "自定义":
            try: target_fps = float(self.fr_val_var.get())
            except ValueError: pass

        for f in self.file_list:
            f_issues = []
            # 检查分辨率
            if target_h and f['height'] < target_h:
                f_issues.append(f"分辨率({f['height']}p < {target_h}p)")
            
            # 检查码率
            if target_bitrate and f['bitrate_kbps'] * 1000 < target_bitrate:
                f_issues.append(f"码率({f['bitrate_kbps']}k < {target_bitrate//1000}k)")
                
            # 检查帧率
            if target_fps and 0 < f['fps'] < target_fps:
                f_issues.append(f"帧率({f['fps']} < {target_fps})")
            
            if f_issues:
                issues.append(f"文件 [{f['filename']}]: " + ", ".join(f_issues))
        
        if issues:
            msg = "以下文件原始参数低于目标值，压缩可能无法改善画质或增大体积:\n\n" + "\n".join(issues[:5])
            if len(issues) > 5: msg += f"\n...等共 {len(issues)} 个文件。"
            msg += "\n\n是否继续处理？"
            return messagebox.askyesno("参数提示", msg)
        
        return True

    def start_compression(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加视频文件。")
            return
        
        if self.is_processing:
            return
        
        # 参数预检
        if not self.check_parameters():
            return

        self.is_processing = True
        self.start_btn.config(state="disabled")
        
        # 准备输出目录
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        self.output_dir = os.path.join(script_dir, OUTPUT_DIR_NAME)
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            
        # 启动后台线程
        threading.Thread(target=self.process_files, daemon=True).start()

    def process_files(self):
        total = len(self.file_list)
        self.log(f"开始处理，共 {total} 个文件...")
        
        for idx, info in enumerate(self.file_list):
            if not self.is_processing: break # 支持停止逻辑扩展
            
            # 更新进度
            self.root.after(0, self.update_progress, idx + 1, total, info['filename'])
            self.root.after(0, self.update_tree_status, info['path'], "压缩中...")
            
            # 构建参数
            params = {
                'codec': self.codec_var.get(),
                'mode': self.mode_var.get(),
                'crf': self.crf_var.get(),
                'bitrate': self.bitrate_var.get(),
                'preset': self.preset_var.get(),
                'resolution_mode': self.res_mode_var.get(),
                'resolution_scale': self.res_scale_var.get(),
                'resolution_custom': self.res_custom_var.get(),
                'framerate_mode': self.fr_mode_var.get(),
                'framerate_custom': self.fr_val_var.get(),
                'audio_bitrate': self.audio_bitrate_var.get()
            }
            
            # 输出路径
            base_name = os.path.splitext(info['filename'])[0]
            output_path = os.path.join(self.output_dir, f"{base_name}_compressed.mp4")
            
            cmd = FFmpegController.build_command(info['path'], output_path, params)
            
            try:
                # 重定向 stdin 防止 ffmpeg 在 Windows 下暂停
                startupinfo = None
                if os.name == 'nt':
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                
                proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo)
                
                if proc.returncode == 0:
                    # 获取新文件信息进行对比
                    new_info = FFmpegController.get_video_info(output_path)
                    if new_info:
                        saved = (1 - new_info['size'] / info['size']) * 100
                        log_msg = (f"完成: {info['filename']}\n"
                                   f"  原始: {info['width']}x{info['height']} @ {info['bitrate_kbps']}k, {info['size']/1024/1024:.2f}MB\n"
                                   f"  新文件: {new_info['width']}x{new_info['height']} @ {new_info['bitrate_kbps']}k, {new_info['size']/1024/1024:.2f}MB\n"
                                   f"  节省空间: {saved:.1f}%")
                        self.root.after(0, self.log, log_msg)
                        self.root.after(0, self.update_tree_status, info['path'], "已完成")
                else:
                    err_msg = proc.stderr.decode('utf-8', errors='ignore')[-200:]
                    self.root.after(0, self.log, f"失败: {info['filename']}. 错误: {err_msg}")
                    self.root.after(0, self.update_tree_status, info['path'], "失败")
                    
            except Exception as e:
                self.root.after(0, self.log, f"处理出错: {info['filename']} - {str(e)}")
                self.root.after(0, self.update_tree_status, info['path'], "错误")

        self.root.after(0, self.finish_compression)

    def update_progress(self, current, total, filename):
        self.progress_var.set(f"进度: {current}/{total} 正在处理: {filename}")
        self.progress_bar['value'] = (current / total) * 100

    def update_tree_status(self, item_id, status):
        try:
            # 更新状态列 (当前状态在第 9 列，索引为 8)
            vals = list(self.tree.item(item_id, 'values'))
            vals[8] = status
            self.tree.item(item_id, values=vals)
        except:
            pass

    def finish_compression(self):
        self.is_processing = False
        self.start_btn.config(state="normal")
        self.progress_var.set("处理完成")
        self.log("所有任务处理完毕。")
        messagebox.showinfo("完成", f"视频压缩完成！\n文件保存在: {self.output_dir}")

if __name__ == "__main__":
    # 检测高DPI支持
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    root = TkinterDnD.Tk()
    app = VideoCompressorApp(root)
    root.mainloop()
