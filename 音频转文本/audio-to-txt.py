# 依赖安装: pip install openai-whisper tkinterdnd2 torch
# 需提前安装 FFmpeg 并配置环境变量

import os
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
import whisper
import torch
import platform
from tkinterdnd2 import TkinterDnD, DND_FILES

class AudioTranscriberApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Whisper 智能音视频转文本工具 Pro Max (增强版)")
        self.root.geometry("1050x750")  # 稍微放大一点以容纳新组件
        self.center_window(1050, 750)
        
        # 路径设置
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_dir = os.path.join(self.script_dir, "转录输出结果")
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)

        self.supported_formats = (
            '.mp3', '.flac', '.wav', '.m4a', '.ogg', '.aac', 
            '.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv'
        )
        
        # 状态变量
        self.model = None
        self.current_model_name = ""
        self.is_processing = False
        self.file_queue = [] 
        
        self.setup_ui()

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        self.root.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

    def get_device_info(self):
        if torch.cuda.is_available():
            dev_name = torch.cuda.get_device_name(0)
            return "GPU 加速", dev_name, "#388E3C" # 绿色
        else:
            proc = platform.processor() or "未知 CPU"
            if len(proc) > 30: proc = proc[:27] + "..."
            return "CPU 慢速", proc, "#D32F2F" # 红色

    def setup_ui(self):
        # --- 顶部设置区 ---
        settings_frame = ttk.LabelFrame(self.root, text="核心设置参数", padding=(10, 8))
        settings_frame.pack(fill=tk.X, padx=10, pady=5)

        # 第1行设置
        ttk.Label(settings_frame, text="🧠 识别模型:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.model_var = tk.StringVar(value="base")
        ttk.Combobox(settings_frame, textvariable=self.model_var, values=["tiny", "base", "small", "medium", "large", "turbo"], width=8, state="readonly").grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(settings_frame, text="🗣️ 音频语言:").grid(row=0, column=2, padx=(15, 5), pady=5, sticky=tk.W)
        self.lang_var = tk.StringVar(value="自动检测")
        ttk.Combobox(settings_frame, textvariable=self.lang_var, values=["自动检测", "中文 (zh)", "英文 (en)", "日文 (ja)"], width=10, state="readonly").grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(settings_frame, text="💾 导出格式:").grid(row=0, column=4, padx=(15, 5), pady=5, sticky=tk.W)
        self.format_var = tk.StringVar(value="纯文本 (.txt)")
        ttk.Combobox(settings_frame, textvariable=self.format_var, values=["纯文本 (.txt)", "字幕文件 (.srt)", "歌词文件 (.lrc)", "全部导出 (TXT+SRT+LRC)"], width=18, state="readonly").grid(row=0, column=5, padx=5, pady=5)

        # 设备提示
        dev_type, dev_name, color = self.get_device_info()
        device_label = tk.Label(settings_frame, text=f"当前设备: [{dev_type}] {dev_name}", fg=color, font=("Microsoft YaHei", 9, "bold"))
        device_label.grid(row=0, column=6, padx=(20, 5), pady=5, sticky=tk.E)

        # 第2行设置：提示词 (解决繁体字问题)
        ttk.Label(settings_frame, text="📝 初始提示词 (Prompt):").grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        self.prompt_var = tk.StringVar(value="简体中文，正确使用标点符号。")
        prompt_entry = ttk.Entry(settings_frame, textvariable=self.prompt_var, width=50)
        prompt_entry.grid(row=1, column=2, columnspan=4, padx=5, pady=5, sticky=tk.W)
        ttk.Label(settings_frame, text="(提示：输入'简体中文'可大幅减少繁体输出)", foreground="#757575").grid(row=1, column=6, sticky=tk.W)

        # --- 中部：左右分栏面板 ---
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 左侧：待处理文件列表
        queue_frame = ttk.LabelFrame(self.paned_window, text="待处理列表 (支持拖拽文件/文件夹到此)", padding=5)
        self.paned_window.add(queue_frame, weight=1)

        list_scroll_frame = tk.Frame(queue_frame)
        list_scroll_frame.pack(fill=tk.BOTH, expand=True)
        
        self.queue_listbox = tk.Listbox(list_scroll_frame, selectmode=tk.EXTENDED, font=("Microsoft YaHei", 10), activestyle="none", bg="#FAFAFA")
        scrollbar_y = ttk.Scrollbar(list_scroll_frame, orient=tk.VERTICAL, command=self.queue_listbox.yview)
        scrollbar_x = ttk.Scrollbar(list_scroll_frame, orient=tk.HORIZONTAL, command=self.queue_listbox.xview)
        self.queue_listbox.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.queue_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        q_btn_frame = tk.Frame(queue_frame)
        q_btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(q_btn_frame, text="📄 添加文件", command=self.select_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(q_btn_frame, text="📁 添加目录", command=self.select_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(q_btn_frame, text="🗑 清空", command=self.clear_queue).pack(side=tk.RIGHT, padx=2)
        ttk.Button(q_btn_frame, text="➖ 删除选中", command=self.remove_selected).pack(side=tk.RIGHT, padx=2)

        # 右侧：处理日志
        log_frame = ttk.LabelFrame(self.paned_window, text="系统处理日志", padding=5)
        self.paned_window.add(log_frame, weight=2) 

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 10), bg="#1E1E1E", fg="#D4D4D4")
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 配置日志颜色标签
        self.log_text.tag_config("success", foreground="#A5D6A7") # 浅绿
        self.log_text.tag_config("error", foreground="#EF9A9A")   # 浅红
        self.log_text.tag_config("highlight", foreground="#90CAF9", font=("Consolas", 10, "bold")) # 浅蓝
        self.log_text.tag_config("warning", foreground="#FFE082") # 浅黄

        log_btn_frame = tk.Frame(log_frame)
        log_btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(log_btn_frame, text="🧹 清空日志", command=self.clear_log).pack(side=tk.RIGHT, padx=2)

        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)

        # --- 底部：执行区与进度条 ---
        action_frame = tk.Frame(self.root)
        action_frame.pack(fill=tk.X, padx=10, pady=(5, 0))

        self.start_btn = tk.Button(
            action_frame, text="▶ 开始批量转录", font=("Microsoft YaHei", 12, "bold"), 
            bg="#2E7D32", fg="white", activebackground="#4CAF50", activeforeground="white",
            command=self.start_transcription, cursor="hand2", pady=8
        )
        self.start_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(action_frame, text="📂 打开输出目录", command=self.open_output_dir).pack(side=tk.RIGHT, ipady=6)

        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, variable=self.progress_var, mode='determinate')
        self.progress_bar.pack(fill=tk.X, padx=10, pady=(5, 0))

        # 状态栏
        self.status_var = tk.StringVar(value="就绪。请添加文件到待处理列表，然后点击【开始批量转录】。")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=(5, 2))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(2, 0))

        self.draw_init_hint()

    def draw_init_hint(self):
        self.clear_log()
        hint = (
            "🚀 欢迎使用 Whisper 音视频转文本工具 Pro Max\n"
            "=================================================\n"
            "💡 操作指南:\n"
            " 1. 将需要处理的音视频拖入左侧【待处理列表】。\n"
            " 2. 调整模型、语言、导出格式，以及初始提示词。\n"
            " 3. 点击下方的绿色按钮【▶ 开始批量转录】。\n\n"
            f"📌 支持的格式: {', '.join(self.supported_formats)}\n"
            "=================================================\n\n"
        )
        self.log(hint, "highlight")

    def log(self, message, tag=None):
        if threading.current_thread() is not threading.main_thread():
            self.root.after(0, self.log, message, tag)
            return
        self.log_text.config(state='normal')
        if tag:
            self.log_text.insert(tk.END, message + "\n", tag)
        else:
            self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def set_status(self, text):
        if threading.current_thread() is not threading.main_thread():
            self.root.after(0, self.set_status, text)
            return
        self.status_var.set(text)

    def update_listbox_item(self, index, text, bg_color=None, fg_color=None):
        """线程安全地更新列表框项"""
        if threading.current_thread() is not threading.main_thread():
            self.root.after(0, self.update_listbox_item, index, text, bg_color, fg_color)
            return
        
        # 防止索引越界
        if index < self.queue_listbox.size():
            self.queue_listbox.delete(index)
            self.queue_listbox.insert(index, text)
            if bg_color: self.queue_listbox.itemconfig(index, {'bg': bg_color})
            if fg_color: self.queue_listbox.itemconfig(index, {'fg': fg_color})

    def open_output_dir(self):
        os.startfile(self.output_dir) if os.name == 'nt' else os.system(f'open "{self.output_dir}"')

    # --- 队列管理 ---
    def select_files(self):
        file_paths = filedialog.askopenfilenames(title="选择音视频文件")
        if file_paths: self.add_to_queue(file_paths)

    def select_folder(self):
        folder_path = filedialog.askdirectory(title="选择包含音视频的文件夹")
        if folder_path: self.add_to_queue([folder_path])

    def on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        if paths: self.add_to_queue(paths)

    def add_to_queue(self, paths):
        if self.is_processing:
            messagebox.showwarning("提示", "正在处理任务中，请等待完成后再添加新文件。")
            return

        added_count = 0
        for path in paths:
            if os.path.isfile(path) and path.lower().endswith(self.supported_formats):
                if path not in self.file_queue:
                    self.file_queue.append(path)
                    self.queue_listbox.insert(tk.END, f"  {os.path.basename(path)}")
                    added_count += 1
            elif os.path.isdir(path):
                for root_dir, _, files in os.walk(path):
                    for file in files:
                        if file.lower().endswith(self.supported_formats):
                            full_path = os.path.join(root_dir, file)
                            if full_path not in self.file_queue:
                                self.file_queue.append(full_path)
                                self.queue_listbox.insert(tk.END, f"  {file}")
                                added_count += 1
        
        if added_count > 0:
            self.log(f"➕ 添加了 {added_count} 个文件。当前共 {len(self.file_queue)} 个待处理。", "highlight")
        else:
            messagebox.showinfo("提示", "未发现支持的新音视频文件。")

    def remove_selected(self):
        if self.is_processing: return
        selected_indices = self.queue_listbox.curselection()
        if not selected_indices: return
        
        for idx in reversed(selected_indices):
            self.queue_listbox.delete(idx)
            del self.file_queue[idx]
        
        self.log(f"➖ 移除了 {len(selected_indices)} 个文件。当前共 {len(self.file_queue)} 个待处理。")

    def clear_queue(self):
        if self.is_processing: return
        self.queue_listbox.delete(0, tk.END)
        self.file_queue.clear()
        self.log("🗑️ 待处理列表已清空。")

    # --- 格式化核心 ---
    def format_timestamp(self, seconds):
        """修复了时间戳精度 Bug，保证毫秒为3位数且不溢出"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        millis = int(round((seconds - int(seconds)) * 1000))
        return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"

    def format_timestamp_lrc(self, seconds):
        minutes = int(seconds // 60)
        secs = int(seconds % 60)
        hundredths = int((seconds - int(seconds)) * 100)
        return f"[{minutes:02d}:{secs:02d}.{hundredths:02d}]"

    def get_safe_output_path(self, base_name, ext):
        """防止同名文件互相覆盖"""
        output_path = os.path.join(self.output_dir, f"{base_name}{ext}")
        counter = 1
        while os.path.exists(output_path):
            output_path = os.path.join(self.output_dir, f"{base_name}_{counter}{ext}")
            counter += 1
        return output_path

    # --- 模型与转录 ---
    def load_model(self):
        target_model = self.model_var.get()
        if self.model is None or self.current_model_name != target_model:
            self.log(f"\n⚙️ 正在加载模型 [{target_model}]，首次运行需下载，请耐心等待...", "warning")
            self.set_status("正在加载AI模型...")
            try:
                self.model = whisper.load_model(target_model)
                self.current_model_name = target_model
                self.log("✅ 模型加载成功！", "success")
            except Exception as e:
                self.log(f"❌ 模型加载失败: {e}", "error")
                return False
        return True

    def start_transcription(self):
        if self.is_processing: return
        if not self.file_queue:
            messagebox.showwarning("提示", "待处理列表为空，请先添加文件！")
            return

        self.is_processing = True
        self.start_btn.config(state=tk.DISABLED, bg="#757575", text="⏳ 正在全力转录中...")
        self.progress_var.set(0)
        self.progress_bar.config(maximum=len(self.file_queue))
        
        current_tasks = self.file_queue.copy()
        threading.Thread(target=self.run_transcription, args=(current_tasks,), daemon=True).start()

    def run_transcription(self, file_list):
        try:
            if not self.load_model():
                self.is_processing = False
                self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL, bg="#2E7D32", text="▶ 开始批量转录"))
                return

            total_files = len(file_list)
            success_count = 0
            
            lang_mapping = {"自动检测": None, "中文 (zh)": "zh", "英文 (en)": "en", "日文 (ja)": "ja"}
            selected_lang = lang_mapping[self.lang_var.get()]
            export_format = self.format_var.get()
            prompt_text = self.prompt_var.get().strip()
            
            use_fp16 = torch.cuda.is_available()

            self.log(f"\n🚀 开始执行批量转录，共 {total_files} 个任务", "highlight")
            self.log("-" * 50)
            
            for i, media_path in enumerate(file_list):
                file_name = os.path.basename(media_path)
                base_name = os.path.splitext(file_name)[0]
                
                # 更新 UI 状态
                self.set_status(f"处理中 ({i+1}/{total_files}): {file_name}")
                self.log(f"⏳ [{i+1}/{total_files}] 正在处理: {file_name}")
                self.update_listbox_item(i, f"▶ [处理中] {file_name}", bg_color="#FFF9C4", fg_color="black")
                
                start_time = time.time()
                
                try:
                    # 组装参数
                    transcribe_opts = {
                        "language": selected_lang,
                        "fp16": use_fp16
                    }
                    if prompt_text:
                        transcribe_opts["initial_prompt"] = prompt_text

                    # 执行核心识别
                    result = self.model.transcribe(media_path, **transcribe_opts)
                    
                    txt_lines, srt_lines, lrc_lines = [], [], []
                    
                    for idx, segment in enumerate(result["segments"]):
                        text = segment['text'].strip()
                        if not text: continue
                        
                        txt_lines.append(text)
                        
                        srt_start = self.format_timestamp(segment['start'])
                        srt_end = self.format_timestamp(segment['end'])
                        srt_lines.append(f"{idx + 1}\n{srt_start} --> {srt_end}\n{text}\n")

                        lrc_time = self.format_timestamp_lrc(segment['start'])
                        lrc_lines.append(f"{lrc_time}{text}")
                    
                    # 保存文件
                    if "纯文本" in export_format or "全部导出" in export_format:
                        safe_path = self.get_safe_output_path(base_name, ".txt")
                        with open(safe_path, "w", encoding="utf-8-sig") as f:
                            f.write('\n'.join(txt_lines))
                            
                    if "字幕文件" in export_format or "全部导出" in export_format:
                        safe_path = self.get_safe_output_path(base_name, ".srt")
                        with open(safe_path, "w", encoding="utf-8-sig") as f:
                            f.write('\n'.join(srt_lines))

                    if "歌词文件" in export_format or "全部导出" in export_format:
                        safe_path = self.get_safe_output_path(base_name, ".lrc")
                        with open(safe_path, "w", encoding="utf-8-sig") as f:
                            f.write('\n'.join(lrc_lines))
                    
                    cost_time = round(time.time() - start_time, 1)
                    self.log(f"   ✅ 完成 (耗时 {cost_time} 秒)", "success")
                    self.update_listbox_item(i, f"✅ [已完成] {file_name}", bg_color="#E8F5E9", fg_color="#2E7D32")
                    success_count += 1

                except Exception as e:
                    self.log(f"   ❌ 失败原因: {e}", "error")
                    self.update_listbox_item(i, f"❌ [失败] {file_name}", bg_color="#FFEBEE", fg_color="#D32F2F")
                
                # 更新进度条
                self.root.after(0, lambda v=i+1: self.progress_var.set(v))
            
            self.log("-" * 50)
            self.log(f"🎉 全部任务结束！成功: {success_count}，失败: {total_files - success_count}", "highlight")
            self.set_status("处理完成！您可以点击右下角按钮打开输出目录。")

        except Exception as e:
            self.log(f"⚠️ 发生严重错误: {e}", "error")
            self.set_status("发生致命错误，请查看日志")
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL, bg="#2E7D32", text="▶ 开始批量转录"))

if __name__ == "__main__":
    try:
        root = TkinterDnD.Tk()
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')
        
        app = AudioTranscriberApp(root)
        root.mainloop()
    except ImportError as e:
        print(f"缺少依赖: {e}")
        print("请在终端运行以下命令安装依赖：")
        print("pip install openai-whisper tkinterdnd2 torch")
