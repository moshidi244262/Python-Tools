# 依赖安装: pip install moviepy tkinterdnd2
# 提示: 如果未安装，程序会自动弹窗提示安装命令

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from typing import List

# --- 兼容性检查与导入 ---
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except ImportError as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("依赖缺失", "请安装缺失的库: tkinterdnd2\n\n请在命令行运行:\npip install tkinterdnd2")
    sys.exit(1)

try:
    # 尝试导入 MoviePy 2.0+
    from moviepy import VideoFileClip
except ImportError:
    try:
        # 兼容导入 MoviePy 1.x
        from moviepy.editor import VideoFileClip
    except ImportError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("依赖缺失", "请安装缺失的库: moviepy\n\n请在命令行运行:\npip install moviepy")
        sys.exit(1)


class VideoToAudioExtractor:
    SUPPORTED_EXTENSIONS = {'.mp4', '.mkv', '.mov', '.avi', '.flv', '.webm', '.wmv', '.m4v', '.mpg', '.mpeg', '.ogv'}
    SUPPORTED_AUDIO_FORMATS = ['mp3', 'wav'] # 可扩展其他格式

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("视频音频提取工具 v2.0 - 优化版")
        self.root.geometry("750x550")
        
        # 默认输出路径设为当前用户桌面
        default_output = os.path.join(os.path.expanduser("~"), "Desktop", "提取的音频")
        self.output_dir = tk.StringVar(value=default_output)
        self.audio_format = tk.StringVar(value="mp3")

        self.file_list = []
        self.is_processing = False

        self._setup_ui()

    def _setup_ui(self):
        """构建现代化的 ttk 界面布局"""
        # --- 顶部设置区 ---
        settings_frame = ttk.LabelFrame(self.root, text=" 输出设置 ", padding=10)
        settings_frame.pack(fill=tk.X, padx=10, pady=5)

        # 路径设置
        ttk.Label(settings_frame, text="保存路径:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(settings_frame, textvariable=self.output_dir, width=55, state='readonly').grid(row=0, column=1, padx=5, pady=5)
        self.btn_change_dir = ttk.Button(settings_frame, text="更改目录", command=self._change_output_dir)
        self.btn_change_dir.grid(row=0, column=2, padx=5, pady=5)

        # 格式设置
        ttk.Label(settings_frame, text="输出格式:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.combo_format = ttk.Combobox(settings_frame, textvariable=self.audio_format, values=self.SUPPORTED_AUDIO_FORMATS, state="readonly", width=10)
        self.combo_format.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        # --- 操作按钮区 ---
        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(fill=tk.X)

        self.btn_select_files = ttk.Button(control_frame, text="➕ 选择文件", command=self._select_files)
        self.btn_select_files.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_select_folder = ttk.Button(control_frame, text="📁 选择文件夹", command=self._select_folder)
        self.btn_select_folder.pack(side=tk.LEFT, padx=5)

        self.btn_clear = ttk.Button(control_frame, text="🗑️ 清空列表", command=self._clear_list)
        self.btn_clear.pack(side=tk.LEFT, padx=5)

        # 使用一个特别颜色的按钮强调“开始”
        self.btn_start = tk.Button(control_frame, text="▶ 开始提取", bg="#4CAF50", fg="white", font=("Microsoft YaHei", 9, "bold"), relief=tk.FLAT, command=self._start_extraction)
        self.btn_start.pack(side=tk.RIGHT, ipadx=10, ipady=2)

        # --- 列表显示区 ---
        list_frame = ttk.LabelFrame(self.root, text=" 待处理文件 (支持拖拽) ", padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, bg="#fafafa", font=("Microsoft YaHei", 9), relief=tk.FLAT, highlightthickness=1)
        self.listbox.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(fill=tk.Y, side=tk.RIGHT)
        self.listbox.config(yscrollcommand=scrollbar.set)

        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self._on_drop)

        # --- 日志与进度区 ---
        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill=tk.X, padx=10, pady=(5, 0))
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=5, state='disabled', font=("Consolas", 9), bg="#f5f5f5")
        self.log_area.pack(fill=tk.X)

        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=10, pady=5)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪：请拖拽文件/文件夹到上方列表。")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=2)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _set_ui_state(self, state):
        """统一控制UI控件的可用状态"""
        self.btn_select_files.config(state=state)
        self.btn_select_folder.config(state=state)
        self.btn_clear.config(state=state)
        self.btn_change_dir.config(state=state)
        self.combo_format.config(state=state)
        
        start_bg = "#4CAF50" if state == tk.NORMAL else "#a5d6a7"
        self.btn_start.config(state=state, bg=start_bg)

    def _log(self, message: str):
        def append():
            self.log_area.config(state='normal')
            self.log_area.insert(tk.END, message + "\n")
            self.log_area.see(tk.END)
            self.log_area.config(state='disabled')
        self.root.after(0, append)

    def _update_status(self, text: str):
        self.root.after(0, lambda: self.status_var.set(text))

    def _update_progress(self, value: float):
        self.root.after(0, lambda: self.progress_var.set(value))

    def _change_output_dir(self):
        new_dir = filedialog.askdirectory(title="选择输出保存位置", initialdir=self.output_dir.get())
        if new_dir:
            self.output_dir.set(new_dir)

    def _add_valid_files(self, paths: List[str]):
        count = 0
        for path_str in paths:
            path = Path(path_str)
            if path.is_file() and path.suffix.lower() in self.SUPPORTED_EXTENSIONS:
                abs_path = str(path.absolute())
                if abs_path not in self.file_list:
                    self.file_list.append(abs_path)
                    self.listbox.insert(tk.END, f"{path.name}  ({abs_path})")
                    count += 1
            elif path.is_dir():
                for ext in self.SUPPORTED_EXTENSIONS:
                    for file in path.rglob(f"*{ext}"):
                        abs_path = str(file.absolute())
                        if abs_path not in self.file_list:
                            self.file_list.append(abs_path)
                            self.listbox.insert(tk.END, f"{file.name}  ({abs_path})")
                            count += 1
                            
        self._update_status(f"已新增 {count} 个文件，当前共 {len(self.file_list)} 个文件。")

    def _select_files(self):
        files = filedialog.askopenfilenames(
            title="选择视频文件",
            filetypes=[("视频文件", "*.mp4 *.mkv *.mov *.avi *.flv *.webm *.wmv *.m4v *.mpg *.mpeg *.ogv"), ("所有文件", "*.*")]
        )
        if files:
            self._add_valid_files(list(files))

    def _select_folder(self):
        folder = filedialog.askdirectory(title="选择包含视频的文件夹")
        if folder:
            self._add_valid_files([folder])

    def _on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        self._add_valid_files(list(files))

    def _clear_list(self):
        self.file_list.clear()
        self.listbox.delete(0, tk.END)
        self._update_progress(0)
        self._update_status("列表已清空。")

    def _start_extraction(self):
        if not self.file_list:
            messagebox.showwarning("提示", "列表为空，请先添加视频文件。")
            return
            
        # 确保输出目录存在
        target_dir = self.output_dir.get()
        try:
            Path(target_dir).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出目录:\n{target_dir}\n错误信息: {e}")
            return

        self.is_processing = True
        self._set_ui_state(tk.DISABLED) # 锁定 UI
        self._log("---------- 开始提取任务 ----------")
        self._update_progress(0)
        threading.Thread(target=self._process_files, daemon=True).start()

    def _process_files(self):
        success_count = 0
        fail_count = 0
        total_files = len(self.file_list)
        target_format = self.audio_format.get()
        target_dir = self.output_dir.get()
        
        for idx, video_path in enumerate(self.file_list, 1):
            try:
                filename = Path(video_path).stem
                output_path = os.path.join(target_dir, f"{filename}.{target_format}")
                
                self._update_status(f"正在处理 ({idx}/{total_files}): {Path(video_path).name}")
                self._log(f"正在提取: {video_path}")

                # 提取逻辑
                with VideoFileClip(video_path) as clip:
                    if clip.audio is None:
                        raise ValueError("该视频不包含音轨")
                    # moviepy 会根据后缀名自动匹配编码器 (mp3, wav 等)
                    clip.audio.write_audiofile(output_path, logger=None) 
                
                self._log(f"成功保存: {output_path}")
                success_count += 1

            except Exception as e:
                self._log(f"失败: {Path(video_path).name} -> {e}")
                fail_count += 1
                
            # 更新进度条
            progress_pct = (idx / total_files) * 100
            self._update_progress(progress_pct)

        # 任务结束回调
        self.root.after(0, lambda: self._on_process_complete(success_count, fail_count, target_dir))

    def _on_process_complete(self, success_count, fail_count, target_dir):
        self._log("---------- 任务结束 ----------")
        self.is_processing = False
        self._set_ui_state(tk.NORMAL) # 解锁 UI
        self._update_status("处理完成。")
        messagebox.showinfo("完成", f"处理完成！\n成功: {success_count} 个\n失败: {fail_count} 个\n\n文件已保存至:\n{target_dir}")


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    
    # 在 Windows 系统上尝试应用原生主题
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    app = VideoToAudioExtractor(root)
    root.mainloop()
