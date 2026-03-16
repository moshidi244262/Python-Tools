# 依赖安装: pip install Pillow tkinterdnd2
import os
import sys
import json
import base64
import threading
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk

# ==========================================
# 依赖检查放在最顶部，防止未安装库直接崩溃
# ==========================================
try:
    from PIL import Image
    from tkinterdnd2 import TkinterDnD, DND_FILES
except ImportError as e:
    # 如果是用python直接运行，弹窗提示
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "缺少依赖", 
        f"运行程序缺少必要的库: {e}\n\n请打开终端(CMD/PowerShell)运行以下命令安装:\n\npip install Pillow tkinterdnd2"
    )
    sys.exit(1)

# ==========================================
# 核心逻辑类
# ==========================================
class CharacterCardExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("酒馆角色卡数据提取工具 v3.0 (增强美化版)")
        self.root.geometry("850x650")
        self.root.minsize(700, 500)
        
        # 初始化默认输出路径 (当前脚本所在目录的子文件夹)
        base_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_dir = tk.StringVar(value=os.path.join(base_dir, "输出_角色数据"))
        
        # 状态变量
        self.is_processing = False
        
        self.setup_ui()
        self.ensure_output_dir()

    def setup_ui(self):
        """配置并初始化现代化的 UI 界面"""
        # 设置 ttk 主题
        style = ttk.Style(self.root)
        # 尝试使用系统中比较现代的主题
        themes = style.theme_names()
        if 'clam' in themes:
            style.theme_use('clam')
        
        # 自定义样式
        style.configure('TButton', font=('Microsoft YaHei', 10), padding=5)
        style.configure('Header.TLabel', font=('Microsoft YaHei', 12, 'bold'))
        style.configure('DropZone.TFrame', background='#e3f2fd')
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 顶部路径设置区
        path_frame = ttk.LabelFrame(main_frame, text=" 💾 输出设置 ", padding=10)
        path_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Entry(path_frame, textvariable=self.output_dir, state='readonly', font=('Consolas', 10)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(path_frame, text="更改目录", command=self.change_output_dir, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(path_frame, text="打开文件夹", command=self.open_output_dir, width=12).pack(side=tk.LEFT, padx=5)

        # 2. 核心操作按钮区
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Button(btn_frame, text="📄 选择文件 (支持多选)", command=self.select_files, width=22).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="📁 选择文件夹 (自动扫描)", command=self.select_folder, width=24).pack(side=tk.LEFT, padx=0)
        ttk.Button(btn_frame, text="🗑️ 清空日志", command=self.clear_log, width=12).pack(side=tk.RIGHT)

        # 3. 拖拽提示区 (美化)
        self.drop_frame = tk.Frame(main_frame, bg="#e1f5fe", highlightbackground="#81d4fa", highlightcolor="#81d4fa", highlightthickness=2, bd=0)
        self.drop_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.drop_label = tk.Label(
            self.drop_frame, 
            text="👇 将 角色卡文件 或 文件夹 拖拽到此处 👇", 
            bg="#e1f5fe", fg="#0277bd", 
            font=("Microsoft YaHei", 14, "bold"), 
            height=3
        )
        self.drop_label.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # 4. 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

        # 5. 日志显示区
        log_frame = ttk.LabelFrame(main_frame, text=" 📝 处理日志 ", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled', font=("Consolas", 10), bg="#f8f9fa", fg="#212529")
        self.log_area.pack(fill=tk.BOTH, expand=True)

        # 6. 底部状态栏
        self.status_var = tk.StringVar(value="准备就绪。请选择文件或直接拖拽入窗口。")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, padding=(10, 2), font=('Microsoft YaHei', 9))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 绑定拖拽事件 (绑定到 Frame 和 Label 以确保都能响应)
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)

    # --- 目录与辅助操作 ---
    def ensure_output_dir(self):
        try:
            target_dir = self.output_dir.get()
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出目录: {e}")

    def change_output_dir(self):
        folder = filedialog.askdirectory(title="选择新的输出文件夹")
        if folder:
            self.output_dir.set(folder)
            self.ensure_output_dir()
            self.log_safe(f"[系统] 输出目录已更改为: {folder}")

    def open_output_dir(self):
        folder = self.output_dir.get()
        if os.path.exists(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("提示", "输出目录目前不存在！")

    def log_safe(self, message):
        """线程安全的日志输出"""
        def _log():
            self.log_area.config(state='normal')
            self.log_area.insert(tk.END, message + "\n")
            self.log_area.see(tk.END)
            self.log_area.config(state='disabled')
        self.root.after(0, _log)

    def set_status_safe(self, message, progress=None):
        """线程安全的状态栏和进度条更新"""
        def _update():
            self.status_var.set(message)
            if progress is not None:
                self.progress_var.set(progress)
        self.root.after(0, _update)

    def clear_log(self):
        self.log_area.config(state='normal')
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state='disabled')

    # --- 交互事件处理 ---
    def select_files(self):
        if self.is_processing: return
        files = filedialog.askopenfilenames(
            title="选择角色卡文件",
            filetypes=[("图片文件", "*.png *.webp *.jpg *.jpeg"), ("所有文件", "*.*")]
        )
        if files: self.start_processing(list(files))

    def select_folder(self):
        if self.is_processing: return
        folder = filedialog.askdirectory(title="选择包含角色卡的文件夹")
        if folder: self.start_processing([folder])

    def on_drop(self, event):
        if self.is_processing: return
        raw_data = event.data
        try:
            paths = self.root.tk.splitlist(raw_data)
        except Exception:
            paths = [raw_data.strip('{}')]
            
        clean_paths = [p.strip('{}') for p in paths]
        if clean_paths:
            self.start_processing(clean_paths)

    # --- 核心处理逻辑 ---
    def start_processing(self, targets):
        self.is_processing = True
        self.progress_var.set(0)
        self.ensure_output_dir()
        
        # 启动后台线程处理
        thread = threading.Thread(target=self.process_targets, args=(targets,))
        thread.daemon = True
        thread.start()

    def process_targets(self, targets):
        all_files = []
        
        self.set_status_safe("正在扫描文件...")
        for path in targets:
            if not path: continue
            if os.path.isfile(path):
                all_files.append(path)
            elif os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for file in files:
                        all_files.append(os.path.join(root, file))
        
        # 过滤图片
        valid_files = [f for f in all_files if f.lower().endswith(('.png', '.webp', '.jpg', '.jpeg'))]
        total = len(valid_files)
        
        if not valid_files:
            self.log_safe("⚠️ 未发现支持的图片文件(.png, .webp, .jpg)。")
            self.set_status_safe("就绪", 0)
            self.is_processing = False
            return

        self.log_safe(f"🚀 开始处理，共发现 {total} 个图片文件...")
        success_count = 0
        fail_count = 0
        out_folder = self.output_dir.get()

        for index, file_path in enumerate(valid_files):
            filename = os.path.basename(file_path)
            try:
                result = self.extract_metadata(file_path)
                if result:
                    # 提取角色名字 (如果JSON里有name字段，优先用name，否则用文件名)
                    chara_name = result.get("data", {}).get("name") or result.get("name", "")
                    
                    base_name = os.path.splitext(filename)[0]
                    # 清理文件名非法字符
                    safe_name = re.sub(r'[\\/*?:"<>|]', "", chara_name if chara_name else base_name).strip()
                    if not safe_name: safe_name = "未命名角色"
                    
                    out_file = os.path.join(out_folder, f"{safe_name}.json")
                    
                    # 查重命名
                    counter = 1
                    while os.path.exists(out_file):
                        out_file = os.path.join(out_folder, f"{safe_name}_{counter}.json")
                        counter += 1
                    
                    with open(out_file, 'w', encoding='utf-8') as f:
                        json.dump(result, f, ensure_ascii=False, indent=4)
                    
                    self.log_safe(f"✅ [成功] {filename} -> {os.path.basename(out_file)}")
                    success_count += 1
                else:
                    self.log_safe(f"⏭️ [跳过] {filename} (未检测到角色数据)")
                    fail_count += 1
            except Exception as e:
                self.log_safe(f"❌ [错误] {filename}: {str(e)}")
                fail_count += 1

            # 更新进度
            current_progress = ((index + 1) / total) * 100
            self.set_status_safe(f"处理中: {index + 1} / {total}", current_progress)

        self.log_safe("-" * 50)
        self.log_safe(f"🎉 处理完成！成功: {success_count} 张, 跳过/失败: {fail_count} 张")
        self.set_status_safe("处理完成", 100)
        self.is_processing = False

    def decode_chara_payload(self, raw_data):
        """尝试将提取出的字符串/字节解码为JSON"""
        try:
            if isinstance(raw_data, bytes):
                # 如果是字节，先尝试去除末尾/开头的杂项如 NULL 字节
                raw_data = raw_data.strip(b'\x00')
                text_content = raw_data.decode('utf-8')
            else:
                text_content = str(raw_data)

            # 1. 如果它已经是明文 JSON
            if text_content.strip().startswith('{'):
                return json.loads(text_content)

            # 2. 尝试 Base64 解码
            # 补齐 Padding
            missing_padding = len(text_content) % 4
            if missing_padding:
                text_content += '=' * (4 - missing_padding)
            
            decoded_bytes = base64.b64decode(text_content)
            return json.loads(decoded_bytes.decode('utf-8'))
        except Exception:
            return None

    def extract_metadata(self, file_path):
        """增强版解析：支持 PNG 文本块、EXIF 以及 底层二进制扫描"""
        # 方法 A: 使用 Pillow 读取标准的 PNG tEXt 块 或 基本 EXIF
        try:
            with Image.open(file_path) as img:
                img.load()
                
                # 1. 检查标准的 PNG 'chara' 字段
                if 'chara' in img.info:
                    parsed = self.decode_chara_payload(img.info['chara'])
                    if parsed: return parsed
                
                # 2. 检查 EXIF 数据 (常用于 WebP 格式的角色卡)
                exif_data = img.getexif()
                if exif_data:
                    # 37510 是 UserComment 的 EXIF Tag ID
                    user_comment = exif_data.get(37510) 
                    if user_comment:
                        # UserComment 有时会带有 'UNICODE\x00' 头
                        if isinstance(user_comment, bytes) and user_comment.startswith(b'UNICODE\x00'):
                            user_comment = user_comment[8:]
                        elif isinstance(user_comment, str) and user_comment.startswith('UNICODE\x00'):
                            user_comment = user_comment[8:]
                            
                        parsed = self.decode_chara_payload(user_comment)
                        if parsed: return parsed

                # 3. 遍历 info 中其他可能的大段文本
                for key, value in img.info.items():
                    if isinstance(value, str) and len(value) > 100:
                        parsed = self.decode_chara_payload(value)
                        if parsed: return parsed
        except Exception:
            pass

        # 方法 B: 暴力二进制特征扫描兜底方案 
        # (处理一些 Pillow 无法正常加载 EXIF 的受损图片或特殊 WebP)
        try:
            with open(file_path, 'rb') as f:
                content = f.read()
                
                # 特征1: 明文 JSON 块 
                # (有些工具会将明文 {"name":"xxx"} 直接存进图片)
                if b'{"spec"' in content or b'{"name"' in content:
                    # 使用正则尝试提取 JSON 结构
                    match = re.search(br'\{.*"name"\s*:.*\}', content, re.DOTALL)
                    if match:
                        try:
                            return json.loads(match.group(0).decode('utf-8'))
                        except: pass

                # 特征2: Base64 特征块 (ey开头是 { 的 Base64)
                # 寻找形如 eyJ 开头且长度足够长的 Base64 字符串
                matches = re.finditer(br'(eyJ[a-zA-Z0-9+/=]{100,})', content)
                for match in matches:
                    parsed = self.decode_chara_payload(match.group(1))
                    if parsed and ("name" in str(parsed) or "data" in str(parsed)):
                        return parsed
        except Exception:
            pass

        return None


if __name__ == "__main__":
    # 使用 TkinterDnD.Tk 代替默认的 tk.Tk
    root = TkinterDnD.Tk()
    
    # 将窗口置于屏幕中央
    root.update_idletasks()
    width = 850
    height = 650
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    app = CharacterCardExtractor(root)
    root.mainloop()
