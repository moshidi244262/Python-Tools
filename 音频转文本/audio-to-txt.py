# 依赖安装: pip install openai-whisper tkinterdnd2
# 注意: 使用本脚本需要提前安装 FFmpeg 并配置环境变量 (如果已能正常运行则无需理会)

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import whisper  # OpenAI Whisper
from tkinterdnd2 import TkinterDnD, DND_FILES

class AudioTranscriberApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Whisper 音频转文本提取工具")
        self.root.geometry("800x600")
        
        # 获取脚本所在目录，并设置输出文件夹路径
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_dir = os.path.join(self.script_dir, "音频文本文件")
        
        # 确保输出文件夹存在
        if not os.path.exists(self.output_dir):
            try:
                os.makedirs(self.output_dir)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录: {e}")
                root.destroy()
                return

        # 支持的音频格式
        self.supported_formats = ('.mp3', '.flac', '.wav')
        
        # 初始化模型变量
        self.model = None
        self.is_processing = False

        # --- 界面布局 ---
        # 顶部控制区
        control_frame = tk.Frame(root, padx=10, pady=10)
        control_frame.pack(fill=tk.X)

        btn_select_files = tk.Button(control_frame, text="选择音频文件", width=15, command=self.select_files)
        btn_select_files.pack(side=tk.LEFT, padx=5)

        btn_select_folder = tk.Button(control_frame, text="选择文件夹", width=15, command=self.select_folder)
        btn_select_folder.pack(side=tk.LEFT, padx=5)

        # 模型选择
        tk.Label(control_frame, text="模型:").pack(side=tk.LEFT, padx=(20, 5))
        self.model_var = tk.StringVar(value="base")
        model_options = ["tiny", "base", "small", "medium", "large"]
        model_menu = tk.OptionMenu(control_frame, self.model_var, *model_options)
        model_menu.pack(side=tk.LEFT)

        # 主显示区 (拖拽区域)
        self.drop_frame = tk.Frame(root, bg="#f0f0f0", relief="sunken", borderwidth=2)
        self.drop_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(self.drop_frame, wrap=tk.WORD, state='disabled', font=("Arial", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 绑定拖拽事件
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        # 也可以绑定到整个窗口，这里绑定到中间frame体验更好

        # 状态栏
        self.status_var = tk.StringVar(value="就绪。请选择文件或拖拽到此区域。")
        status_bar = tk.Label(root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 绘制提示文字
        self.draw_drop_hint()

    def draw_drop_hint(self):
        """在日志区绘制初始提示（仅在空白时有效）"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, "操作说明:\n1. 点击按钮选择单个或多个音频文件\n2. 点击按钮选择文件夹（含子文件夹）\n3. 直接拖拽音频文件或文件夹到此区域\n\n支持的格式: mp3, flac, wav\n输出目录: 脚本所在目录下的'音频文本文件'文件夹\n")
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

    def set_status(self, text):
        """更新状态栏"""
        self.status_var.set(text)

    def load_model(self):
        """懒加载模型"""
        if self.model is None:
            model_name = self.model_var.get()
            self.log(f"正在加载 Whisper 模型 ({model_name})，首次加载需要下载，请稍候...")
            try:
                # device="cuda" if torch.cuda.is_available() else "cpu" 可以优化，但whisper会自动处理
                self.model = whisper.load_model(model_name)
                self.log("模型加载完成。")
            except Exception as e:
                self.log(f"模型加载失败: {e}")
                return False
        return True

    def select_files(self):
        """选择文件对话框"""
        file_paths = filedialog.askopenfilenames(
            title="选择音频文件",
            filetypes=[("音频文件", "*.mp3 *.flac *.wav"), ("所有文件", "*.*")]
        )
        if file_paths:
            self.process_paths(file_paths)

    def select_folder(self):
        """选择文件夹对话框"""
        folder_path = filedialog.askdirectory(title="选择包含音频的文件夹")
        if folder_path:
            self.process_paths([folder_path])

    def on_drop(self, event):
        """拖拽事件处理"""
        # [修复点] 使用 Tkinter 自带的 splitlist 方法，完美处理带空格和不带空格的多个路径
        paths = self.root.tk.splitlist(event.data)
        
        if paths:
            self.process_paths(paths)

    def process_paths(self, paths):
        """处理输入的路径列表（文件或文件夹）"""
        if self.is_processing:
            messagebox.showwarning("提示", "当前有任务正在处理，请稍候。")
            return

        # 收集所有待处理的音频文件
        audio_files = []
        for path in paths:
            if os.path.isfile(path):
                if path.lower().endswith(self.supported_formats):
                    audio_files.append(path)
            elif os.path.isdir(path):
                # 递归查找音频文件
                for root_dir, _, files in os.walk(path):
                    for file in files:
                        if file.lower().endswith(self.supported_formats):
                            audio_files.append(os.path.join(root_dir, file))
        
        if not audio_files:
            messagebox.showwarning("提示", "未找到支持的音频文件。")
            return

        # 清空日志并开始处理
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END) # 清空旧日志
        self.log_text.config(state='disabled')
        
        # 启动后台线程处理
        self.is_processing = True
        threading.Thread(target=self.run_transcription, args=(audio_files,), daemon=True).start()

    def run_transcription(self, file_list):
        """后台处理线程"""
        try:
            if not self.load_model():
                self.is_processing = False
                return

            total_files = len(file_list)
            success_count = 0
            
            self.log(f"开始处理 {total_files} 个音频文件...\n")
            
            for i, audio_path in enumerate(file_list):
                self.set_status(f"处理中 ({i+1}/{total_files}): {os.path.basename(audio_path)}")
                self.log(f"[{i+1}/{total_files}] 正在识别: {os.path.basename(audio_path)}")
                
                try:
                    # 调用 Whisper 进行识别
                    # task="transcribe" 是默认值，language="zh" 可以强制中文，但自动检测效果通常也不错
                    result = self.model.transcribe(audio_path, fp16=False) 
                    
                    # 提取文本并处理分行
                    # Whisper 的 segments 包含时间戳，我们利用它来进行合理的换行
                    lines = []
                    
                    for segment in result["segments"]:
                        text = segment['text'].strip()
                        # 简单的分行逻辑：依据 Whisper 切分的 segment，每个 segment 为新的一行
                        if text:
                            lines.append(text)
                    
                    final_text = '\n'.join(lines)
                    
                    # 构造输出路径
                    relative_name = os.path.basename(audio_path)
                    txt_name = os.path.splitext(relative_name)[0] + ".txt"
                    save_path = os.path.join(self.output_dir, txt_name)
                    
                    # 写入文件 (使用 utf-8-sig 防止 Windows 记事本乱码)
                    with open(save_path, "w", encoding="utf-8-sig") as f:
                        f.write(final_text)
                    
                    self.log(f"    -> 完成: {txt_name}")
                    success_count += 1

                except Exception as e:
                    self.log(f"    -> 错误: 处理失败。原因: {e}")
            
            self.log("\n" + "="*30)
            self.log(f"所有任务完成。成功: {success_count}, 失败: {total_files - success_count}")
            self.log(f"文件已保存在: {self.output_dir}")
            self.set_status("处理完成")

        except Exception as e:
            self.log(f"发生严重错误: {e}")
            self.set_status("发生错误")
        finally:
            self.is_processing = False

if __name__ == "__main__":
    # 检查 ffmpeg 是否存在 (Whisper依赖)
    # whisper 只在调用时检查，这里提前做一个简单的提示
    try:
        # 使用 TkinterDnD 创建主窗口
        root = TkinterDnD.Tk()
        app = AudioTranscriberApp(root)
        root.mainloop()
    except ImportError as e:
        print(f"缺少依赖库: {e}")
        print("请运行: pip install openai-whisper tkinterdnd2")
    except Exception as e:
        print(f"启动失败: {e}")
