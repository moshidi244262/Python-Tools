# 依赖安装: pip install openai-whisper tkinterdnd2 torch
# 需提前安装 FFmpeg 并配置环境变量

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk  # 引入 ttk 以美化界面
import whisper
import torch
import platform
from tkinterdnd2 import TkinterDnD, DND_FILES
import datetime

class AudioTranscriberApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Whisper 智能音视频转文本工具 Pro Max")
        self.root.geometry("1000x700")  # 加宽窗口以容纳左右双面板
        
        # 居中显示窗口
        self.center_window(1000, 700)
        
        # 路径设置
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_dir = os.path.join(self.script_dir, "转录输出结果")
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)

        # 扩展支持的格式 (Whisper/FFmpeg 支持视频和常见音频)
        self.supported_formats = (
            '.mp3', '.flac', '.wav', '.m4a', '.ogg', '.aac', 
            '.mp4', '.mkv', '.avi', '.mov'
        )
        
        # 状态变量
        self.model = None
        self.current_model_name = ""
        self.is_processing = False
        self.file_queue = []  # 用于存储待处理文件的绝对路径

        self.setup_ui()

    def center_window(self, width, height):
        """将窗口居中显示"""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        self.root.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

    def get_device_info(self):
        """获取真实的硬件名称"""
        if torch.cuda.is_available():
            dev_name = torch.cuda.get_device_name(0)
            return "GPU (CUDA)", dev_name, "green"
        else:
            proc = platform.processor()
            if not proc:
                proc = "未知 CPU (可能为 Apple Silicon / 基础 x86)"
            # 如果名称太长则截断以防破坏 UI
            if len(proc) > 35:
                proc = proc[:32] + "..."
            return "CPU", proc, "red"

    def setup_ui(self):
        """构建美化后的 UI"""
        # --- 顶部设置区 ---
        settings_frame = ttk.LabelFrame(self.root, text="核心设置", padding=(10, 5))
        settings_frame.pack(fill=tk.X, padx=10, pady=5)

        # 模型选择
        ttk.Label(settings_frame, text="识别模型:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.model_var = tk.StringVar(value="base")
        model_combo = ttk.Combobox(settings_frame, textvariable=self.model_var, values=["tiny", "base", "small", "medium", "large", "turbo"], width=8, state="readonly")
        model_combo.grid(row=0, column=1, padx=5, pady=5)

        # 语言选择
        ttk.Label(settings_frame, text="音频语言:").grid(row=0, column=2, padx=(10, 5), pady=5, sticky=tk.W)
        self.lang_var = tk.StringVar(value="自动检测")
        lang_combo = ttk.Combobox(settings_frame, textvariable=self.lang_var, values=["自动检测", "中文 (zh)", "英文 (en)", "日文 (ja)"], width=10, state="readonly")
        lang_combo.grid(row=0, column=3, padx=5, pady=5)

        # 导出格式选择
        ttk.Label(settings_frame, text="导出格式:").grid(row=0, column=4, padx=(10, 5), pady=5, sticky=tk.W)
        self.format_var = tk.StringVar(value="纯文本 (.txt)")
        format_combo = ttk.Combobox(settings_frame, textvariable=self.format_var, values=["纯文本 (.txt)", "字幕文件 (.srt)", "歌词文件 (.lrc)", "全部导出 (TXT+SRT+LRC)"], width=18, state="readonly")
        format_combo.grid(row=0, column=5, padx=5, pady=5)

        # 计算设备提示
        dev_type, dev_name, color = self.get_device_info()
        device_str = f"当前设备: [{dev_type}] {dev_name}"
        device_label = ttk.Label(settings_frame, text=device_str, foreground=color)
        device_label.grid(row=0, column=6, padx=(15, 5), pady=5, sticky=tk.E)

        # --- 中部：左右分栏面板 ---
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # ---------------- 左侧：待处理文件列表 ----------------
        queue_frame = ttk.LabelFrame(self.paned_window, text="待处理列表 (支持拖拽文件/文件夹到此)", padding=5)
        self.paned_window.add(queue_frame, weight=1)

        # 列表框与滚动条
        list_scroll_frame = tk.Frame(queue_frame)
        list_scroll_frame.pack(fill=tk.BOTH, expand=True)
        
        self.queue_listbox = tk.Listbox(list_scroll_frame, selectmode=tk.EXTENDED, font=("Microsoft YaHei", 9), activestyle="none")
        scrollbar_y = ttk.Scrollbar(list_scroll_frame, orient=tk.VERTICAL, command=self.queue_listbox.yview)
        scrollbar_x = ttk.Scrollbar(list_scroll_frame, orient=tk.HORIZONTAL, command=self.queue_listbox.xview)
        self.queue_listbox.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.queue_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 列表操作按钮
        q_btn_frame = tk.Frame(queue_frame)
        q_btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(q_btn_frame, text="📄 添加文件", command=self.select_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(q_btn_frame, text="📁 添加目录", command=self.select_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(q_btn_frame, text="🗑 清空列表", command=self.clear_queue).pack(side=tk.RIGHT, padx=2)
        ttk.Button(q_btn_frame, text="➖ 删除选中", command=self.remove_selected).pack(side=tk.RIGHT, padx=2)

        # ---------------- 右侧：处理日志 ----------------
        log_frame = ttk.LabelFrame(self.paned_window, text="处理日志", padding=5)
        self.paned_window.add(log_frame, weight=2)  # 给日志区分配多一倍的宽度

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 10), bg="#1e1e1e", fg="#d4d4d4")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        log_btn_frame = tk.Frame(log_frame)
        log_btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(log_btn_frame, text="🧹 清空日志", command=self.clear_log).pack(side=tk.RIGHT, padx=2)

        # 注册全局拖拽事件
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)

        # --- 底部：执行区与状态栏 ---
        action_frame = tk.Frame(self.root)
        action_frame.pack(fill=tk.X, padx=10, pady=5)

        # 显眼的开始按钮
        self.start_btn = tk.Button(
            action_frame, text="▶ 开始批量转录", font=("Microsoft YaHei", 12, "bold"), 
            bg="#2E7D32", fg="white", activebackground="#4CAF50", activeforeground="white",
            command=self.start_transcription, cursor="hand2", pady=5
        )
        self.start_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        ttk.Button(action_frame, text="📂 打开输出目录", command=self.open_output_dir).pack(side=tk.RIGHT, ipady=4)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪。请添加文件到待处理列表，然后点击【开始批量转录】。")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=(5, 2))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.draw_init_hint()

    def draw_init_hint(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        hint = (
            "🚀 欢迎使用 Whisper 音视频转文本工具 Pro Max\n\n"
            "操作指南:\n"
            "1. 将需要处理的音视频拖入左侧【待处理列表】，可以任意删除不想处理的文件。\n"
            "2. 在上方调整需要的模型、语言和输出格式。\n"
            "3. 确认无误后，点击下方的绿色按钮【▶ 开始批量转录】。\n"
            f"支持的格式: {', '.join(self.supported_formats)}\n\n"
        )
        self.log_text.insert(tk.END, hint)
        self.log_text.config(state='disabled')

    def log(self, message):
        """线程安全的日志输出"""
        if threading.current_thread() is not threading.main_thread():
            self.root.after(0, self.log, message)
            return
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def clear_log(self):
        """清空日志区域"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        self.log("🧹 日志已清空。")

    def set_status(self, text):
        if threading.current_thread() is not threading.main_thread():
            self.root.after(0, self.set_status, text)
            return
        self.status_var.set(text)

    def open_output_dir(self):
        """打开输出文件夹"""
        os.startfile(self.output_dir) if os.name == 'nt' else os.system(f'open "{self.output_dir}"')

    # --- 队列管理 ---
    def select_files(self):
        file_paths = filedialog.askopenfilenames(title="选择音视频文件")
        if file_paths:
            self.add_to_queue(file_paths)

    def select_folder(self):
        folder_path = filedialog.askdirectory(title="选择包含音视频的文件夹")
        if folder_path:
            self.add_to_queue([folder_path])

    def on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        if paths:
            self.add_to_queue(paths)

    def add_to_queue(self, paths):
        """将选中的文件/文件夹解析并加入待处理列表"""
        if self.is_processing:
            messagebox.showwarning("提示", "正在处理任务中，请等待完成后再添加新文件。")
            return

        added_count = 0
        for path in paths:
            if os.path.isfile(path):
                if path.lower().endswith(self.supported_formats):
                    if path not in self.file_queue:
                        self.file_queue.append(path)
                        self.queue_listbox.insert(tk.END, os.path.basename(path))
                        added_count += 1
            elif os.path.isdir(path):
                for root_dir, _, files in os.walk(path):
                    for file in files:
                        if file.lower().endswith(self.supported_formats):
                            full_path = os.path.join(root_dir, file)
                            if full_path not in self.file_queue:
                                self.file_queue.append(full_path)
                                self.queue_listbox.insert(tk.END, file)
                                added_count += 1
        
        if added_count > 0:
            self.log(f"➕ 成功添加 {added_count} 个文件到待处理列表。当前共 {len(self.file_queue)} 个文件。")
        else:
            messagebox.showinfo("提示", "未发现支持的新音视频文件。")

    def remove_selected(self):
        """删除列表框中选中的项目"""
        if self.is_processing:
            return
        selected_indices = self.queue_listbox.curselection()
        if not selected_indices:
            return
        
        # 必须从后往前删，否则索引会错乱
        for idx in reversed(selected_indices):
            self.queue_listbox.delete(idx)
            del self.file_queue[idx]
        
        self.log(f"➖ 已从列表中移除 {len(selected_indices)} 个文件。当前共 {len(self.file_queue)} 个文件。")

    def clear_queue(self):
        """清空整个待处理列表"""
        if self.is_processing:
            return
        self.queue_listbox.delete(0, tk.END)
        self.file_queue.clear()
        self.log("🗑️ 待处理列表已清空。")

    # --- 格式化与模型加载 ---
    def format_timestamp(self, seconds):
        td = datetime.timedelta(seconds=seconds)
        hours, remainder = divmod(td.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        milliseconds = int(td.microseconds / 1000)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d},{milliseconds:03d}"

    def format_timestamp_lrc(self, seconds):
        minutes = int(seconds // 60)
        secs = int(seconds % 60)
        hundredths = int((seconds - int(seconds)) * 100)
        return f"[{minutes:02d}:{secs:02d}.{hundredths:02d}]"

    def load_model(self):
        target_model = self.model_var.get()
        if self.model is None or self.current_model_name != target_model:
            self.log(f"\n⚙️ 正在加载/切换模型至 [{target_model}]，请稍候...")
            self.set_status("正在加载AI模型...")
            try:
                self.model = whisper.load_model(target_model)
                self.current_model_name = target_model
                self.log("✅ 模型加载成功！")
            except Exception as e:
                self.log(f"❌ 模型加载失败: {e}")
                return False
        return True

    # --- 核心转录业务 ---
    def start_transcription(self):
        """用户点击【开始转录】按钮后触发"""
        if self.is_processing:
            return
            
        if not self.file_queue:
            messagebox.showwarning("提示", "待处理列表为空，请先添加文件！")
            return

        self.is_processing = True
        # 禁用按钮，防止连按
        self.start_btn.config(state=tk.DISABLED, bg="#757575", text="⏳ 正在转录处理中...")
        
        # 拷贝一份队列进行处理，防止多线程环境下列表被意外修改
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
            
            use_fp16 = torch.cuda.is_available()

            self.log(f"▶ 开始处理，本次共执行 {total_files} 个任务")
            self.log(f"▶ 输出目录: {self.output_dir}\n" + "-"*40)
            
            for i, media_path in enumerate(file_list):
                file_name = os.path.basename(media_path)
                base_name = os.path.splitext(file_name)[0]
                self.set_status(f"处理中 ({i+1}/{total_files}): {file_name}")
                self.log(f"⏳ [{i+1}/{total_files}] 正在转录: {file_name}")
                
                try:
                    result = self.model.transcribe(
                        media_path, 
                        language=selected_lang,
                        fp16=use_fp16
                    )
                    
                    txt_lines = []
                    srt_lines = []
                    lrc_lines = []
                    
                    for idx, segment in enumerate(result["segments"]):
                        text = segment['text'].strip()
                        if not text: continue
                        
                        txt_lines.append(text)
                        
                        start_time = self.format_timestamp(segment['start'])
                        end_time = self.format_timestamp(segment['end'])
                        srt_lines.append(f"{idx + 1}\n{start_time} --> {end_time}\n{text}\n")

                        lrc_time = self.format_timestamp_lrc(segment['start'])
                        lrc_lines.append(f"{lrc_time}{text}")
                    
                    if "纯文本" in export_format or "全部导出" in export_format:
                        with open(os.path.join(self.output_dir, base_name + ".txt"), "w", encoding="utf-8-sig") as f:
                            f.write('\n'.join(txt_lines))
                            
                    if "字幕文件" in export_format or "全部导出" in export_format:
                        with open(os.path.join(self.output_dir, base_name + ".srt"), "w", encoding="utf-8-sig") as f:
                            f.write('\n'.join(srt_lines))

                    if "歌词文件" in export_format or "全部导出" in export_format:
                        with open(os.path.join(self.output_dir, base_name + ".lrc"), "w", encoding="utf-8-sig") as f:
                            f.write('\n'.join(lrc_lines))
                    
                    self.log(f"   ✅ 完成")
                    success_count += 1

                except Exception as e:
                    self.log(f"   ❌ 失败原因: {e}")
            
            self.log("-" * 40)
            self.log(f"🎉 全部任务结束！成功: {success_count}，失败: {total_files - success_count}")
            self.set_status("处理完成")

        except Exception as e:
            self.log(f"⚠️ 发生严重错误: {e}")
            self.set_status("发生错误")
        finally:
            self.is_processing = False
            # 恢复按钮状态
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL, bg="#2E7D32", text="▶ 开始批量转录"))

if __name__ == "__main__":
    try:
        root = TkinterDnD.Tk()
        style = ttk.Style()
        style.theme_use('clam')
        
        app = AudioTranscriberApp(root)
        root.mainloop()
    except ImportError as e:
        print(f"缺少依赖: {e}")
        print("请在终端运行以下命令安装依赖：")
        print("pip install openai-whisper tkinterdnd2 torch")
