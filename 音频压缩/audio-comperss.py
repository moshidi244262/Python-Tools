# 依赖安装: pip install mutagen tkinterdnd2
# 终极优化版：移除 pydub 依赖，直接使用原生 subprocess 调用 ffmpeg，完美兼容 Python 3.13/3.14+
# 注意：需要安装ffmpeg并添加到系统PATH，或将ffmpeg.exe放在脚本同目录

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import traceback
from pathlib import Path
import subprocess
import shutil

# ==========================================
# 安全导入第三方库 (防止双击由于缺环境导致闪退)
# ==========================================
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    import mutagen
    from mutagen.flac import FLAC
    from mutagen.wave import WAVE
    from mutagen.mp3 import MP3, EasyMP3
    from mutagen.aiff import AIFF
except ImportError as e:
    # 捕获导入错误并使用系统弹窗提示
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    error_msg = f"启动失败，缺少必要的第三方库：\n{str(e)}\n\n" \
                f"当前使用的Python路径：\n{sys.executable}\n\n" \
                f"请确保运行以下命令安装环境（已移除pydub依赖）：\n" \
                f"pip install mutagen tkinterdnd2"
    messagebox.showerror("环境配置错误", error_msg)
    sys.exit(1)


class AudioCompressorApp:
    """音频压缩工具主类"""
    
    SUPPORTED_FORMATS = ('.flac', '.wav', '.mp3')
    
    def __init__(self, root):
        self.root = root
        self.root.title("音频压缩工具 - FLAC/WAV/MP3 转 MP3 (原生FFmpeg驱动)")
        self.root.geometry("1200x800")
        self.root.minsize(900, 600)
        
        # 数据存储
        self.audio_files = {}  # {filepath: info_dict}
        self.output_dir = Path(__file__).parent / "压缩音频"
        
        # 压缩参数变量
        self.bitrate_var = tk.StringVar(value="192")
        self.sample_rate_var = tk.StringVar(value="44100")
        self.channels_var = tk.StringVar(value="2")
        
        # 标志位
        self.is_compressing = False
        
        self._setup_ui()
        self._setup_styles()
        
    def _setup_styles(self):
        """设置样式"""
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Microsoft YaHei', 10, 'bold'))
        style.configure('Info.TLabel', font=('Microsoft YaHei', 9))
        
    def _setup_ui(self):
        """构建用户界面"""
        # ========== 顶部参数设置区域 ==========
        param_frame = ttk.LabelFrame(self.root, text="压缩参数设置", padding=10)
        param_frame.pack(fill='x', padx=10, pady=5)
        
        # 第一行参数
        row1 = ttk.Frame(param_frame)
        row1.pack(fill='x', pady=2)
        
        # 比特率
        ttk.Label(row1, text="比特率:", width=10).pack(side='left')
        self.bitrate_combo = ttk.Combobox(
            row1, textvariable=self.bitrate_var, 
            values=["64", "96", "128", "160", "192", "224", "256", "320"],
            width=10, state='readonly'
        )
        self.bitrate_combo.pack(side='left', padx=(0, 20))
        ttk.Label(row1, text="kbps", width=6).pack(side='left')
        
        # 采样率
        ttk.Label(row1, text="采样率:", width=10).pack(side='left')
        self.sample_rate_combo = ttk.Combobox(
            row1, textvariable=self.sample_rate_var,
            values=["8000", "11025", "16000", "22050", "32000", "44100", "48000", "96000"],
            width=10, state='readonly'
        )
        self.sample_rate_combo.pack(side='left', padx=(0, 20))
        ttk.Label(row1, text="Hz", width=6).pack(side='left')
        
        # 声道数
        ttk.Label(row1, text="声道数:", width=10).pack(side='left')
        self.channels_combo = ttk.Combobox(
            row1, textvariable=self.channels_var,
            values=["1 (单声道)", "2 (立体声)"],
            width=12, state='readonly'
        )
        self.channels_combo.pack(side='left', padx=(0, 20))
        
        # 输出格式提示
        ttk.Label(row1, text="输出格式: MP3", style='Title.TLabel').pack(side='right', padx=10)
        
        # ========== 操作按钮区域 ==========
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill='x', padx=10, pady=5)
        
        # 左侧操作按钮
        left_btn_frame = ttk.Frame(btn_frame)
        left_btn_frame.pack(side='left')
        
        ttk.Button(left_btn_frame, text="📁 选择文件", command=self._select_files, width=12).pack(side='left', padx=3)
        ttk.Button(left_btn_frame, text="📂 选择文件夹", command=self._select_folder, width=12).pack(side='left', padx=3)
        ttk.Button(left_btn_frame, text="🗑 清空列表", command=self._clear_list, width=12).pack(side='left', padx=3)
        ttk.Button(left_btn_frame, text="❌ 删除选中", command=self._delete_selected, width=12).pack(side='left', padx=3)
        ttk.Button(left_btn_frame, text="📝 清空日志", command=self._clear_log, width=12).pack(side='left', padx=3)
        
        # 右侧压缩按钮
        self.compress_btn = ttk.Button(btn_frame, text="🚀 开始压缩", command=self._start_compress, width=15)
        self.compress_btn.pack(side='right', padx=3)
        
        # ========== 文件列表区域 ==========
        list_frame = ttk.LabelFrame(self.root, text="音频文件列表 (支持拖拽上传文件或文件夹)", padding=5)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # 创建Treeview
        columns = ('filename', 'format', 'bitrate', 'sample_rate', 'bits', 'channels', 'frame_width', 'size', 'status')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', selectmode='extended', height=10)
        
        # 设置列标题和宽度
        column_config = {
            'filename': ('文件名', 280),
            'format': ('格式', 60),
            'bitrate': ('比特率', 100),
            'sample_rate': ('采样率', 100),
            'bits': ('采样宽度', 80),
            'channels': ('声道数', 60),
            'frame_width': ('帧宽', 80),
            'size': ('文件大小', 100),
            'status': ('状态', 150)
        }
        
        for col, (title, width) in column_config.items():
            self.tree.heading(col, text=title, anchor='center')
            self.tree.column(col, width=width, anchor='center', minwidth=50)
        
        # 滚动条
        scrollbar_y = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(list_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # 布局
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar_y.pack(side='right', fill='y')
        
        # 拖拽绑定
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self._on_drop)
        self.tree.dnd_bind('<<DragEnter>>', lambda e: self.tree.focus())
        
        # ========== 日志区域 ==========
        log_frame = ttk.LabelFrame(self.root, text="压缩日志", padding=5)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.log_text = tk.Text(log_frame, height=12, wrap='word', font=('Consolas', 9))
        log_scrollbar = ttk.Scrollbar(log_frame, orient='vertical', command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scrollbar.pack(side='right', fill='y')
        
        # 配置日志标签颜色
        self.log_text.tag_configure('info', foreground='#2196F3')
        self.log_text.tag_configure('success', foreground='#4CAF50')
        self.log_text.tag_configure('warning', foreground='#FF9800')
        self.log_text.tag_configure('error', foreground='#F44336')
        self.log_text.tag_configure('compare', foreground='#9C27B0')
        
    def _log(self, message, level='info'):
        """添加日志消息"""
        self.log_text.insert('end', message + '\n', level)
        self.log_text.see('end')
        
    def _clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, 'end')
        
    def _format_size(self, size_bytes):
        """格式化文件大小"""
        try:
            size_bytes = float(size_bytes)
            for unit in ['B', 'KB', 'MB', 'GB']:
                if size_bytes < 1024:
                    return f"{size_bytes:.2f} {unit}"
                size_bytes /= 1024
            return f"{size_bytes:.2f} TB"
        except:
            return "N/A"
            
    def _get_audio_info(self, filepath):
        """获取音频文件详细信息"""
        try:
            info = {
                'filepath': filepath,
                'filename': os.path.basename(filepath),
                'format': os.path.splitext(filepath)[1][1:].upper(),
                'bitrate': 'N/A',
                'sample_rate': 'N/A',
                'bits_per_sample': 'N/A',
                'channels': 'N/A',
                'frame_width': 'N/A',
                'size': self._format_size(os.path.getsize(filepath)),
                'size_bytes': os.path.getsize(filepath)
            }
            
            # 使用mutagen获取元数据
            try:
                audio = mutagen.File(filepath, easy=True)
                if audio is not None:
                    audio_info = audio.info
                    
                    # 比特率
                    if hasattr(audio_info, 'bitrate') and audio_info.bitrate:
                        info['bitrate'] = f"{int(audio_info.bitrate / 1000)} kbps"
                        info['bitrate_value'] = audio_info.bitrate
                    
                    # 采样率
                    if hasattr(audio_info, 'sample_rate') and audio_info.sample_rate:
                        info['sample_rate'] = f"{int(audio_info.sample_rate)} Hz"
                        info['sample_rate_value'] = audio_info.sample_rate
                    
                    # 声道数
                    if hasattr(audio_info, 'channels'):
                        info['channels'] = str(audio_info.channels)
                        info['channels_value'] = audio_info.channels
                    
                    # 采样宽度 (bits per sample)
                    if hasattr(audio_info, 'bits_per_sample') and audio_info.bits_per_sample:
                        info['bits_per_sample'] = f"{audio_info.bits_per_sample} bit"
                        info['bits_value'] = audio_info.bits_per_sample
                    elif hasattr(audio_info, 'sample_size') and audio_info.sample_size:
                        info['bits_per_sample'] = f"{audio_info.sample_size} bit"
                        info['bits_value'] = audio_info.sample_size
                    
                    # 帧宽
                    if hasattr(audio_info, 'frame_width'):
                        info['frame_width'] = f"{audio_info.frame_width} bytes"
                    elif info.get('channels_value') and info.get('bits_value'):
                        frame_width = info['channels_value'] * info['bits_value'] // 8
                        info['frame_width'] = f"{frame_width} bytes"
                        
            except Exception as e:
                self._log(f"读取元数据警告 [{os.path.basename(filepath)}]: {str(e)}", 'warning')
            
            return info
            
        except Exception as e:
            self._log(f"获取音频信息失败 [{filepath}]: {str(e)}", 'error')
            return None
            
    def _add_file(self, filepath):
        """添加单个文件"""
        try:
            filepath = os.path.abspath(filepath)
            
            # 检查文件是否存在
            if not os.path.isfile(filepath):
                return False
                
            # 检查是否已存在
            if filepath in self.audio_files:
                return False
                
            # 检查格式
            ext = os.path.splitext(filepath)[1].lower()
            if ext not in self.SUPPORTED_FORMATS:
                return False
            
            # 获取音频信息
            info = self._get_audio_info(filepath)
            if info:
                self.audio_files[filepath] = info
                item_id = self.tree.insert('', 'end', values=(
                    info['filename'],
                    info['format'],
                    info['bitrate'],
                    info['sample_rate'],
                    info['bits_per_sample'],
                    info['channels'],
                    info['frame_width'],
                    info['size'],
                    '待压缩'
                ))
                info['tree_id'] = item_id
                return True
            return False
            
        except Exception as e:
            self._log(f"添加文件失败 [{filepath}]: {str(e)}", 'error')
            return False
            
    def _add_files(self, filepaths):
        """批量添加文件"""
        count = 0
        for filepath in filepaths:
            if self._add_file(filepath):
                count += 1
        self._log(f"已添加 {count} 个音频文件", 'info')
        
    def _scan_folder(self, folder_path):
        """递归扫描文件夹中的音频文件"""
        count = 0
        try:
            for root, dirs, files in os.walk(folder_path):
                for filename in files:
                    ext = os.path.splitext(filename)[1].lower()
                    if ext in self.SUPPORTED_FORMATS:
                        filepath = os.path.join(root, filename)
                        if self._add_file(filepath):
                            count += 1
        except Exception as e:
            self._log(f"扫描文件夹失败: {str(e)}", 'error')
        
        return count
        
    def _select_files(self):
        """选择文件对话框"""
        filetypes = [
            ("支持的音频文件", "*.flac *.wav *.mp3"),
            ("FLAC 文件", "*.flac"),
            ("WAV 文件", "*.wav"),
            ("MP3 文件", "*.mp3"),
            ("所有文件", "*.*")
        ]
        filepaths = filedialog.askopenfilenames(filetypes=filetypes)
        if filepaths:
            self._add_files(filepaths)
            
    def _select_folder(self):
        """选择文件夹对话框"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            count = self._scan_folder(folder_path)
            self._log(f"从文件夹中扫描到 {count} 个音频文件", 'info')
            
    def _on_drop(self, event):
        """处理拖拽事件"""
        try:
            # 解析拖拽的数据
            data = event.data
            # 处理Windows路径格式
            paths = []
            
            # 尝试使用tkinter内置方法分割
            try:
                paths = list(self.root.tk.splitlist(data))
            except:
                # 手动解析
                import re
                paths = re.findall(r'\{([^}]+)\}|(\S+)', data)
                paths = [p[0] or p[1] for p in paths if p[0] or p[1]]
            
            file_count = 0
            folder_count = 0
            
            for path in paths:
                # 清理路径
                path = path.strip().strip('{}').strip('"\'')
                
                if os.path.isfile(path):
                    ext = os.path.splitext(path)[1].lower()
                    if ext in self.SUPPORTED_FORMATS:
                        if self._add_file(path):
                            file_count += 1
                elif os.path.isdir(path):
                    count = self._scan_folder(path)
                    if count > 0:
                        folder_count += 1
                        file_count += count
                        
            if file_count > 0:
                msg = f"拖拽添加了 {file_count} 个音频文件"
                if folder_count > 0:
                    msg += f" (来自 {folder_count} 个文件夹)"
                self._log(msg, 'info')
                
        except Exception as e:
            self._log(f"拖拽处理错误: {str(e)}", 'error')
            traceback.print_exc()
            
    def _clear_list(self):
        """清空文件列表"""
        if self.audio_files:
            if messagebox.askyesno("确认", "确定要清空文件列表吗？"):
                self.audio_files.clear()
                for item in self.tree.get_children():
                    self.tree.delete(item)
                self._log("已清空文件列表", 'info')
                
    def _delete_selected(self):
        """删除选中的文件"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择要删除的文件")
            return
            
        count = len(selected_items)
        if messagebox.askyesno("确认", f"确定要从列表中删除选中的 {count} 个文件吗？"):
            for item_id in selected_items:
                # 查找并删除对应的项目
                for filepath, info in list(self.audio_files.items()):
                    if info.get('tree_id') == item_id:
                        del self.audio_files[filepath]
                        break
                self.tree.delete(item_id)
            self._log(f"已删除 {count} 个选中文件", 'info')
            
    def _check_parameters(self):
        """检查压缩参数，返回是否继续"""
        user_sr = int(self.sample_rate_var.get())
        user_br = int(self.bitrate_var.get()) * 1000  # 转换为bps
        user_ch = int(self.channels_var.get()[0])  # 取数字
        
        warnings = []
        
        for filepath, info in self.audio_files.items():
            issues = []
            
            # 检查采样率
            orig_sr = info.get('sample_rate_value', 0)
            if orig_sr and orig_sr < user_sr:
                issues.append(f"原采样率({int(orig_sr)}Hz) < 设定值({user_sr}Hz)")
            
            # 检查比特率
            orig_br = info.get('bitrate_value', 0)
            if orig_br and orig_br < user_br:
                issues.append(f"原比特率约({int(orig_br/1000)}kbps) < 设定值({int(user_br/1000)}kbps)")
            
            # 检查声道
            orig_ch = info.get('channels_value', 0)
            if orig_ch and orig_ch < user_ch:
                issues.append(f"原声道数({orig_ch}) < 设定值({user_ch})")
                
            if issues:
                warnings.append(f"【{info['filename']}】\n  - " + "\n  - ".join(issues))
        
        if warnings:
            warning_msg = "以下音频的原始参数低于设定值:\n\n" + "\n\n".join(warnings[:5])
            if len(warnings) > 5:
                warning_msg += f"\n\n... 还有 {len(warnings)-5} 个文件有类似问题"
            warning_msg += "\n\n是否继续压缩？(将保持原参数或降低参数处理)"
            
            return messagebox.askyesno("参数警告", warning_msg)
            
        return True
        
    def _start_compress(self):
        """开始压缩"""
        if not self.audio_files:
            messagebox.showwarning("提示", "请先添加要压缩的音频文件")
            return
            
        if self.is_compressing:
            messagebox.showwarning("提示", "正在压缩中，请稍候...")
            return
            
        # 检查参数
        if not self._check_parameters():
            return
            
        # 创建输出目录
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出目录: {str(e)}")
            return
            
        # 禁用按钮，启动压缩线程
        self.is_compressing = True
        self.compress_btn.configure(state='disabled')
        
        thread = threading.Thread(target=self._compress_all, daemon=True)
        thread.start()
        
    def _compress_all(self):
        """压缩所有文件（线程中运行）"""
        total = len(self.audio_files)
        success_count = 0
        fail_count = 0
        
        self._log(f"{'='*60}", 'info')
        self._log(f"开始压缩，共 {total} 个文件", 'info')
        self._log(f"目标参数: 比特率={self.bitrate_var.get()}kbps, 采样率={self.sample_rate_var.get()}Hz, 声道={self.channels_var.get()}", 'info')
        self._log(f"输出目录: {self.output_dir}", 'info')
        self._log(f"{'='*60}", 'info')
        
        for idx, (filepath, info) in enumerate(list(self.audio_files.items()), 1):
            try:
                # 更新状态
                tree_id = info.get('tree_id')
                if tree_id:
                    self.tree.item(tree_id, values=(
                        info['filename'], info['format'], info['bitrate'],
                        info['sample_rate'], info['bits_per_sample'], info['channels'],
                        info['frame_width'], info['size'], f"压缩中 ({idx}/{total})"
                    ))
                
                self._log(f"\n[{idx}/{total}] 正在压缩: {info['filename']}", 'info')
                
                # 压缩文件
                result = self._compress_file(filepath, info)
                
                if result:
                    success_count += 1
                    if tree_id:
                        self.tree.item(tree_id, values=(
                            info['filename'], info['format'], info['bitrate'],
                            info['sample_rate'], info['bits_per_sample'], info['channels'],
                            info['frame_width'], info['size'], '✓ 压缩成功'
                        ), tags=('success',))
                else:
                    fail_count += 1
                    if tree_id:
                        self.tree.item(tree_id, values=(
                            info['filename'], info['format'], info['bitrate'],
                            info['sample_rate'], info['bits_per_sample'], info['channels'],
                            info['frame_width'], info['size'], '✗ 压缩失败'
                        ), tags=('error',))
                        
            except Exception as e:
                fail_count += 1
                self._log(f"压缩失败: {str(e)}", 'error')
                traceback.print_exc()
        
        # 完成
        self._log(f"\n{'='*60}", 'info')
        self._log(f"压缩完成！成功: {success_count}, 失败: {fail_count}", 'success' if fail_count == 0 else 'warning')
        self._log(f"输出位置: {self.output_dir}", 'info')
        
        # 恢复按钮状态
        self.is_compressing = False
        self.root.after(0, lambda: self.compress_btn.configure(state='normal'))
        
    def _compress_file(self, filepath, orig_info):
        """使用 FFmpeg 原生命令压缩单个文件"""
        try:
            self._log(f"  准备处理...", 'info')
            
            # 获取原始参数 (从已解析的 mutagen 信息中拿)
            orig_sr = orig_info.get('sample_rate_value', 44100)
            orig_ch = orig_info.get('channels_value', 2)
            orig_sw_bits = orig_info.get('bits_value', 16)
            orig_size = orig_info.get('size_bytes', 0)
            
            # 获取用户设置参数
            user_sr = int(self.sample_rate_var.get())
            user_br = self.bitrate_var.get()
            user_ch = int(self.channels_var.get()[0])
            
            # 智能调整参数（不提升原始参数，防止体积反向增加）
            target_sr = min(orig_sr, user_sr) if orig_sr else user_sr
            target_ch = min(orig_ch, user_ch) if orig_ch else user_ch
            
            self._log(f"  原始参数: {orig_sr}Hz, {orig_ch}声道, {orig_sw_bits}bit, {self._format_size(orig_size)}", 'info')
            self._log(f"  目标参数: {target_sr}Hz, {target_ch}声道, {user_br}kbps", 'info')
            
            # 生成输出文件名
            base_name = os.path.splitext(orig_info['filename'])[0]
            output_name = f"{base_name}.mp3"
            output_path = self.output_dir / output_name
            
            # 处理重名
            counter = 1
            while output_path.exists():
                output_name = f"{base_name}_{counter}.mp3"
                output_path = self.output_dir / output_name
                counter += 1
            
            self._log(f"  正在调用 FFmpeg 编码MP3...", 'info')
            
            # 构建FFmpeg命令
            ffmpeg_cmd = [
                "ffmpeg",
                "-y",                   # 覆盖输出文件
                "-i", str(filepath),    # 输入文件
                "-ar", str(target_sr),  # 采样率
                "-ac", str(target_ch),  # 声道数
                "-b:a", f"{user_br}k",  # 比特率 (恒定比特率CBR)
                "-map_metadata", "0",   # 保留原始元数据标签
                str(output_path)        # 输出文件
            ]
            
            # 隐藏 Windows 弹出的黑色 CMD 窗口
            creationflags = 0
            if sys.platform == "win32":
                creationflags = subprocess.CREATE_NO_WINDOW
                
            # 执行压缩
            process = subprocess.run(
                ffmpeg_cmd, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE, 
                text=True,
                creationflags=creationflags
            )
            
            # 检查 FFmpeg 执行结果
            if process.returncode != 0:
                self._log(f"  ✗ FFmpeg 错误日志: {process.stderr}", 'error')
                return False
            
            # 获取压缩后文件信息
            new_size = os.path.getsize(output_path)
            new_info = self._get_audio_info(str(output_path))
            
            # 计算压缩率
            compression_ratio = (1 - new_size / orig_size) * 100 if orig_size > 0 else 0
            
            # 显示对比信息
            self._log(f"  " + "─"*50, 'compare')
            self._log(f"  │ 项目       │ 原始值           │ 压缩后           │", 'compare')
            self._log(f"  ├────────────┼──────────────────┼──────────────────│", 'compare')
            self._log(f"  │ 文件大小   │ {self._format_size(orig_size):>16} │ {self._format_size(new_size):>16} │", 'compare')
            self._log(f"  │ 采样率     │ {orig_sr:>14} Hz │ {new_info.get('sample_rate', 'N/A'):>16} │", 'compare')
            self._log(f"  │ 声道数     │ {orig_ch:>16} │ {new_info.get('channels', 'N/A'):>16} │", 'compare')
            self._log(f"  │ 比特率     │ {orig_info.get('bitrate', 'N/A'):>16} │ {new_info.get('bitrate', 'N/A'):>16} │", 'compare')
            self._log(f"  │ 采样宽度   │ {orig_sw_bits:>14} bit │ {new_info.get('bits_per_sample', 'N/A'):>16} │", 'compare')
            self._log(f"  └────────────┴──────────────────┴──────────────────┘", 'compare')
            self._log(f"  压缩率: {compression_ratio:.1f}%", 'success' if compression_ratio > 0 else 'warning')
            self._log(f"  输出文件: {output_name}", 'success')
            
            return True
            
        except Exception as e:
            self._log(f"  ✗ 压缩失败: {str(e)}", 'error')
            traceback.print_exc()
            return False


def main():
    """主函数"""
    try:
        # 创建主窗口
        root = TkinterDnD.Tk()
        
        # 设置窗口图标（如果存在）
        try:
            root.iconbitmap(default='')
        except:
            pass
            
        # 检查ffmpeg并进行可视化提醒
        if shutil.which("ffmpeg") is None:
            messagebox.showwarning(
                "缺少 ffmpeg 组件", 
                "未检测到 ffmpeg 工具！\n\n【必须配置】程序依赖 FFmpeg 进行音频压缩。\n\n解决办法：\n请确保 ffmpeg.exe 放在本脚本同级目录下，或将其添加到系统环境变量 PATH 中。"
            )
        
        # 创建应用
        app = AudioCompressorApp(root)
        
        # 居中显示窗口
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        # 运行主循环
        root.mainloop()
        
    except Exception as e:
        # 如果是 tkdnd 注册失败等运行时错误，通过弹窗显示
        temp_root = tk.Tk()
        temp_root.withdraw()
        messagebox.showerror("程序崩溃", f"运行时发生严重错误：\n{str(e)}\n\n请检查tkinterdnd2是否正确安装。")
        sys.exit(1)


if __name__ == '__main__':
    main()
