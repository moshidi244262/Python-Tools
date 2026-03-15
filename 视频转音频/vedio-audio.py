# 依赖安装: pip install moviepy tkinterdnd2
# 注意: moviepy 依赖 imageio-ffmpeg，首次运行可能需要下载 ffmpeg 二进制文件

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
from typing import List

# 检查第三方库是否安装
try:
    # 【修改点】：适配 MoviePy 2.0+，去掉了 .editor
    from moviepy import VideoFileClip
    from tkinterdnd2 import TkinterDnD, DND_FILES
except ImportError as e:
    root = tk.Tk()
    root.withdraw()
    missing_lib = str(e).split("named")[1].strip() if "named" in str(e) else str(e)
    messagebox.showerror("依赖缺失", f"请安装缺失的库: {missing_lib}\n\n请在命令行运行:\npip install moviepy tkinterdnd2")
    sys.exit(1)

class VideoToAudioExtractor:
    # 支持的视频格式
    SUPPORTED_EXTENSIONS = {'.mp4', '.mkv', '.mov', '.avi', '.flv', '.webm', '.wmv', '.m4v', '.mpg', '.mpeg', '.ogv'}
    # 目标保存路径
    OUTPUT_DIR = r"C:\Users\24426\Desktop\Py工具\提取视频中的音频\提取的音频"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("视频音频提取工具 v1.0")
        self.root.geometry("700x500")
        
        # 确保输出目录存在
        try:
            Path(self.OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出目录:\n{self.OUTPUT_DIR}\n错误信息: {e}")
            sys.exit(1)

        self.file_list = []  # 存储待处理的文件路径
        self.is_processing = False  # 线程锁标志

        self._setup_ui()

    def _setup_ui(self):
        """构建界面布局"""
        # 顶部控制区
        control_frame = tk.Frame(self.root, pady=10)
        control_frame.pack(fill=tk.X)

        btn_select_files = tk.Button(control_frame, text="选择视频文件", width=15, command=self._select_files)
        btn_select_files.pack(side=tk.LEFT, padx=5)

        btn_select_folder = tk.Button(control_frame, text="选择文件夹", width=15, command=self._select_folder)
        btn_select_folder.pack(side=tk.LEFT, padx=5)

        btn_start = tk.Button(control_frame, text="开始提取", width=15, bg="#4CAF50", fg="white", command=self._start_extraction)
        btn_start.pack(side=tk.LEFT, padx=5)

        btn_clear = tk.Button(control_frame, text="清空列表", width=15, command=self._clear_list)
        btn_clear.pack(side=tk.LEFT, padx=5)

        # 文件显示区（支持拖拽）
        list_frame = tk.Frame(self.root)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

        # 使用 tkinterdnd2 实现拖拽
        self.listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, bg="#f0f0f0", font=("Arial", 10))
        self.listbox.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(fill=tk.Y, side=tk.RIGHT)
        self.listbox.config(yscrollcommand=scrollbar.set)

        # 绑定拖拽事件
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self._on_drop)

        # 底部状态栏
        self.status_var = tk.StringVar(value="就绪：请拖拽文件/文件夹到此处，或使用上方按钮选择。")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 日志输出区
        self.log_area = scrolledtext.ScrolledText(self.root, height=6, state='disabled', font=("Consolas", 9))
        self.log_area.pack(fill=tk.X, padx=10, pady=(0, 5))

    def _log(self, message: str):
        """线程安全的日志输出"""
        def append():
            self.log_area.config(state='normal')
            self.log_area.insert(tk.END, message + "\n")
            self.log_area.see(tk.END)
            self.log_area.config(state='disabled')
        
        if threading.current_thread() is threading.main_thread():
            append()
        else:
            self.root.after(0, append)

    def _update_status(self, text: str):
        self.root.after(0, lambda: self.status_var.set(text))

    def _add_valid_files(self, paths: List[str]):
        """筛选并添加有效的视频文件"""
        count = 0
        for path_str in paths:
            path = Path(path_str)
            # 如果是文件，检查后缀
            if path.is_file():
                if path.suffix.lower() in self.SUPPORTED_EXTENSIONS:
                    abs_path = str(path.absolute())
                    if abs_path not in self.file_list:
                        self.file_list.append(abs_path)
                        self.listbox.insert(tk.END, path.name)
                        count += 1
            # 如果是文件夹，递归查找
            elif path.is_dir():
                for ext in self.SUPPORTED_EXTENSIONS:
                    for file in path.rglob(f"*{ext}"):
                        abs_path = str(file.absolute())
                        if abs_path not in self.file_list:
                            self.file_list.append(abs_path)
                            self.listbox.insert(tk.END, str(file.relative_to(path.parent)))
                            count += 1
        self._update_status(f"已添加 {count} 个视频文件，当前列表共 {len(self.file_list)} 个文件。")

    def _select_files(self):
        """处理按钮：选择文件"""
        files = filedialog.askopenfilenames(
            title="选择视频文件",
            filetypes=[("Video Files", "*.mp4 *.mkv *.mov *.avi *.flv *.webm *.wmv *.m4v *.mpg *.mpeg *.ogv"), ("All Files", "*.*")]
        )
        if files:
            self._add_valid_files(list(files))

    def _select_folder(self):
        """处理按钮：选择文件夹"""
        folder = filedialog.askdirectory(title="选择包含视频的文件夹")
        if folder:
            self._add_valid_files([folder])

    def _on_drop(self, event):
        """处理拖拽事件"""
        # TkinterDnD 返回的是类似 {C:/path1} {C:/path2} 的字符串
        # 使用 tk.splitlist 处理包含空格的路径
        files = self.root.tk.splitlist(event.data)
        self._add_valid_files(list(files))

    def _clear_list(self):
        """清空列表"""
        self.file_list.clear()
        self.listbox.delete(0, tk.END)
        self._update_status("列表已清空。")

    def _start_extraction(self):
        """启动提取线程"""
        if not self.file_list:
            messagebox.showwarning("提示", "列表为空，请先添加视频文件。")
            return
        
        if self.is_processing:
            messagebox.showinfo("提示", "正在处理中，请稍候...")
            return

        self.is_processing = True
        self._log("---------- 开始提取任务 ----------")
        threading.Thread(target=self._process_files, daemon=True).start()

    def _process_files(self):
        """后台处理逻辑"""
        success_count = 0
        fail_count = 0
        
        for idx, video_path in enumerate(self.file_list, 1):
            try:
                filename = Path(video_path).stem
                output_path = os.path.join(self.OUTPUT_DIR, f"{filename}.mp3")
                
                self._update_status(f"正在处理 ({idx}/{len(self.file_list)}): {Path(video_path).name}")
                self._log(f"正在提取: {video_path}")

                # 核心提取逻辑
                with VideoFileClip(video_path) as clip:
                    if clip.audio is None:
                        raise ValueError("该视频不包含音轨")
                    
                    clip.audio.write_audiofile(output_path, logger=None) # logger=None 禁用 moviepy 的控制台输出
                
                self._log(f"成功保存: {output_path}")
                success_count += 1

            except Exception as e:
                self._log(f"失败: {Path(video_path).name} -> {e}")
                fail_count += 1

        # 任务结束回调
        self.root.after(0, lambda: messagebox.showinfo("完成", f"处理完成！\n成功: {success_count}\n失败: {fail_count}\n\n文件已保存至:\n{self.OUTPUT_DIR}"))
        self._log("---------- 任务结束 ----------")
        self.is_processing = False
        self._update_status("处理完成。")

if __name__ == "__main__":
    # 使用 TkinterDnD.Tk() 初始化主窗口以支持拖拽
    root = TkinterDnD.Tk()
    app = VideoToAudioExtractor(root)
    root.mainloop()
