# 依赖安装: pip install Pillow tkinterdnd2

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
from datetime import datetime
import re

# ================= 依赖检测与导入 =================
# 尝试导入 Pillow (如果未安装，弹出友好的窗口提示而不是直接闪退)
try:
    from PIL import Image
except ImportError:
    # 创建一个隐藏的主窗口用于弹窗
    err_root = tk.Tk()
    err_root.withdraw()
    messagebox.showerror(
        "缺少运行依赖", 
        "运行此程序需要安装 Pillow 库！\n\n请打开 CMD 命令行或终端，输入以下命令安装：\npip install Pillow"
    )
    sys.exit(1)

# 尝试导入 tkinterdnd2 (拖拽功能)
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False
# =================================================

class ImageCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片压缩工具 v2.0 - 支持格式转换")
        self.root.geometry("1150x780")
        self.root.minsize(950, 650)
        
        # 数据存储
        self.image_data = {}
        self.item_counter = 0
        
        # 路径设置
        self.script_dir = Path(__file__).parent if "__file__" in dir() else Path.cwd()
        self.output_dir = self.script_dir / "压缩图片"
        
        # 支持的输入格式
        self.supported_formats = {'.jpg', '.jpeg', '.png', '.bmp', '.webp', '.tiff', '.gif', '.ico'}
        
        # 构建界面
        self.setup_ui()
        
        # 绑定拖拽
        if HAS_DND:
            self.setup_dnd()
        
    def setup_ui(self):
        """构建用户界面"""
        # 主布局容器
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # === 左侧：文件列表区域 ===
        left_frame = ttk.Frame(main_paned)
        main_paned.add(left_frame, weight=3)
        
        # 工具栏
        toolbar = ttk.Frame(left_frame)
        toolbar.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Button(toolbar, text="📁 选择文件", command=self.select_files, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="📂 选择文件夹", command=self.select_folder, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Button(toolbar, text="🗑 删除选中", command=self.delete_selected, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🧹 清空列表", command=self.clear_all, width=10).pack(side=tk.LEFT, padx=2)
        
        # 提示
        tip = ttk.Label(left_frame, text="💡 提示：支持拖拽文件或文件夹到下方列表 | PNG转JPG/WebP可大幅减小体积", foreground="gray")
        tip.pack(anchor=tk.W, pady=(0, 2))
        
        # 列表 Treeview
        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        cols = ("文件名", "原始大小", "输出格式", "压缩比率", "压缩后大小", "压缩率", "状态")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended")
        
        # 定义列
        self.tree.heading("文件名", text="文件名")
        self.tree.heading("原始大小", text="原始大小")
        self.tree.heading("输出格式", text="输出格式")
        self.tree.heading("压缩比率", text="质量/Q")
        self.tree.heading("压缩后大小", text="压缩后大小")
        self.tree.heading("压缩率", text="压缩率")
        self.tree.heading("状态", text="状态")
        
        self.tree.column("文件名", width=220, anchor=tk.W)
        self.tree.column("原始大小", width=85, anchor=tk.CENTER)
        self.tree.column("输出格式", width=80, anchor=tk.CENTER)
        self.tree.column("压缩比率", width=70, anchor=tk.CENTER)
        self.tree.column("压缩后大小", width=90, anchor=tk.CENTER)
        self.tree.column("压缩率", width=70, anchor=tk.CENTER)
        self.tree.column("状态", width=80, anchor=tk.CENTER)
        
        # 滚动条
        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定事件
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        if HAS_DND:
            self.tree.drop_target_register(DND_FILES)
            
        # === 右侧：设置面板 ===
        right_frame = ttk.Frame(main_paned, width=300)
        main_paned.add(right_frame, weight=1)
        
        # 1. 输出设置
        output_settings = ttk.LabelFrame(right_frame, text="输出设置", padding=10)
        output_settings.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(output_settings, text="输出格式:").pack(anchor=tk.W)
        self.output_format = tk.StringVar(value="保持原格式")
        format_combo = ttk.Combobox(output_settings, textvariable=self.output_format, state="readonly")
        format_combo['values'] = ("保持原格式", "JPEG (.jpg)", "WebP (.webp)", "PNG (.png)", "ICO (.ico)", "BMP (.bmp)")
        format_combo.pack(fill=tk.X, pady=(2, 5))
        format_combo.bind("<<ComboboxSelected>>", self.on_format_change)
        
        # 格式说明
        self.format_hint = ttk.Label(output_settings, text="当前：智能保持原格式", foreground="gray", wraplength=250)
        self.format_hint.pack(anchor=tk.W)
        
        # 2. 全局压缩设置
        global_frame = ttk.LabelFrame(right_frame, text="全局压缩质量", padding=10)
        global_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(global_frame, text="质量/压缩等级 (1%-99%):").pack(anchor=tk.W)
        
        scale_frame = ttk.Frame(global_frame)
        scale_frame.pack(fill=tk.X, pady=5)
        
        self.global_ratio = tk.IntVar(value=80)
        self.global_scale = ttk.Scale(scale_frame, from_=1, to=99, variable=self.global_ratio, orient=tk.HORIZONTAL)
        self.global_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.global_label = ttk.Label(scale_frame, text="80%", width=5, anchor=tk.CENTER)
        self.global_label.pack(side=tk.RIGHT)
        self.global_scale.configure(command=self.update_global_label)
        
        # 快捷按钮
        quick_frame = ttk.Frame(global_frame)
        quick_frame.pack(fill=tk.X, pady=5)
        for r in [50, 70, 80, 90, 95]:
            ttk.Button(quick_frame, text=f"{r}%", width=5,
                       command=lambda val=r: self.set_global_ratio(val)).pack(side=tk.LEFT, padx=1)
        
        ttk.Button(global_frame, text="应用到所有图片", command=self.apply_global_settings).pack(fill=tk.X, pady=(5, 0))
        
        # 3. 自定义压缩设置
        custom_frame = ttk.LabelFrame(right_frame, text="自定义设置 (选中图片)", padding=10)
        custom_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 自定义格式
        ttk.Label(custom_frame, text="输出格式:").pack(anchor=tk.W)
        self.custom_format = tk.StringVar(value="保持原格式")
        self.custom_format_combo = ttk.Combobox(custom_frame, textvariable=self.custom_format, state="readonly")
        self.custom_format_combo['values'] = ("保持原格式", "JPEG (.jpg)", "WebP (.webp)", "PNG (.png)", "ICO (.ico)", "BMP (.bmp)")
        self.custom_format_combo.pack(fill=tk.X, pady=(2, 5))
        
        # 自定义质量
        ttk.Label(custom_frame, text="压缩质量:").pack(anchor=tk.W)
        custom_scale_frame = ttk.Frame(custom_frame)
        custom_scale_frame.pack(fill=tk.X, pady=5)
        
        self.custom_ratio = tk.IntVar(value=80)
        self.custom_scale = ttk.Scale(custom_scale_frame, from_=1, to=99, variable=self.custom_ratio, orient=tk.HORIZONTAL)
        self.custom_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.custom_label = ttk.Label(custom_scale_frame, text="80%", width=5, anchor=tk.CENTER)
        self.custom_label.pack(side=tk.RIGHT)
        self.custom_scale.configure(command=self.update_custom_label)
        
        ttk.Button(custom_frame, text="应用到选中图片", command=self.apply_custom_settings).pack(fill=tk.X, pady=(5, 0))
        
        # 4. 操作区
        action_frame = ttk.LabelFrame(right_frame, text="操作", padding=10)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.compress_btn = ttk.Button(action_frame, text="🚀 开始压缩", command=self.start_compress)
        self.compress_btn.pack(fill=tk.X, pady=2)
        
        ttk.Button(action_frame, text="📂 打开输出目录", command=self.open_output_dir).pack(fill=tk.X, pady=2)
        ttk.Button(action_frame, text="🗑 清空日志", command=self.clear_log).pack(fill=tk.X, pady=2)
        
        # 统计
        self.stats_var = tk.StringVar(value="共 0 张图片")
        ttk.Label(action_frame, textvariable=self.stats_var).pack(anchor=tk.W, pady=(5, 0))
        
        # === 底部：日志和进度 ===
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        log_label = ttk.LabelFrame(bottom_frame, text="运行日志", padding=5)
        log_label.pack(fill=tk.X)
        
        self.log_text = tk.Text(log_label, height=5, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(log_label, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 进度条
        prog_frame = ttk.Frame(self.root)
        prog_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(prog_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, side=tk.LEFT, expand=True)
        self.progress_label = ttk.Label(prog_frame, text="")
        self.progress_label.pack(side=tk.RIGHT, padx=(10, 0))

    def setup_dnd(self):
        """设置拖拽功能"""
        def on_drop(event):
            paths = self.parse_drop_data(event.data)
            self.add_images_from_paths(paths)
        
        def on_drag_enter(event):
            self.tree.configure(background="#f0f8ff")
            
        def on_drag_leave(event):
            self.tree.configure(background="white")
            
        self.tree.dnd_bind('<<Drop>>', on_drop)
        self.tree.dnd_bind('<<DragEnter>>', on_drag_enter)
        self.tree.dnd_bind('<<DragLeave>>', on_drag_leave)
    
    def parse_drop_data(self, data):
        """解析拖拽数据"""
        # 处理 Windows 的拖拽数据格式
        pattern = re.compile(r'\{([^}]+)\}|(\S+)')
        matches = pattern.findall(data)
        paths = []
        for group in matches:
            path = group[0] if group[0] else group[1]
            if path:
                paths.append(path)
        return paths

    def log(self, msg, level="INFO"):
        """写入日志"""
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{ts}][{level}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        
    def clear_log(self):
        """清空日志"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)
        
    def fmt_size(self, b):
        """格式化文件大小"""
        if b < 1024: return f"{b} B"
        if b < 1024**2: return f"{b/1024:.1f} KB"
        return f"{b/1024**2:.2f} MB"
    
    def update_global_label(self, val=None):
        self.global_label.config(text=f"{self.global_ratio.get()}%")
        
    def update_custom_label(self, val=None):
        self.custom_label.config(text=f"{self.custom_ratio.get()}%")
        
    def set_global_ratio(self, val):
        self.global_ratio.set(val)
        self.update_global_label()
        
    def on_format_change(self, event=None):
        """全局格式变更提示"""
        fmt = self.output_format.get()
        hints = {
            "保持原格式": "智能保持原格式",
            "JPEG (.jpg)": "适合照片，体积小，不支持透明",
            "WebP (.webp)": "新一代格式，体积最小，支持透明",
            "PNG (.png)": "无损压缩，适合图标/插画，体积较大",
            "ICO (.ico)": "图标格式，支持多尺寸，体积中等",
            "BMP (.bmp)": "无压缩位图，体积巨大"
        }
        self.format_hint.config(text=f"提示: {hints.get(fmt, '')}")
        
    def on_tree_select(self, event):
        """列表选择事件，同步右侧设置"""
        sels = self.tree.selection()
        if len(sels) == 1:
            item_id = sels[0]
            if item_id in self.image_data:
                data = self.image_data[item_id]
                self.custom_ratio.set(data["ratio"].get())
                self.custom_format.set(data["format"].get())
                self.update_custom_label()
                
    def select_files(self):
        files = filedialog.askopenfilenames(title="选择图片", filetypes=[("图片", "*.jpg *.png *.jpeg *.bmp *.webp *.gif *.ico"), ("所有", "*.*")])
        if files: self.add_images_from_paths(files)
        
    def select_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder: self.add_images_from_paths([folder])
        
    def add_images_from_paths(self, paths):
        """批量添加图片"""
        count = 0
        for p in paths:
            path = Path(p)
            if not path.exists(): continue
            
            if path.is_file():
                if self.add_single_image(path): count += 1
            elif path.is_dir():
                count += self.scan_folder(path)
                
        if count > 0:
            self.log(f"添加了 {count} 张图片")
            self.update_stats()
        else:
            self.log("未找到有效图片", "WARN")
            
    def scan_folder(self, folder):
        """递归扫描文件夹"""
        count = 0
        for root, _, files in os.walk(folder):
            for f in files:
                if self.add_single_image(Path(root) / f): count += 1
        return count
    
    def add_single_image(self, path):
        """添加单张图片"""
        if path.suffix.lower() not in self.supported_formats:
            return False
        
        path_str = str(path.resolve())
        # 去重
        for d in self.image_data.values():
            if d["path"] == path_str: return False
            
        try:
            size = path.stat().st_size
            item_id = f"img_{self.item_counter}"
            self.item_counter += 1
            
            # 初始设置跟随全局
            ratio_var = tk.IntVar(value=self.global_ratio.get())
            format_var = tk.StringVar(value=self.output_format.get())
            
            self.tree.insert("", tk.END, iid=item_id, values=(
                path.name, self.fmt_size(size), 
                format_var.get().split(" ")[0], 
                f"{ratio_var.get()}%", "-", "-", "待处理"
            ))
            
            self.image_data[item_id] = {
                "path": path_str, "name": path.name,
                "size": size, "ratio": ratio_var, "format": format_var
            }
            return True
        except Exception as e:
            self.log(f"添加失败 {path.name}: {e}", "ERROR")
            return False
            
    def delete_selected(self):
        """删除选中项"""
        sels = self.tree.selection()
        if not sels:
            messagebox.showwarning("提示", "请先选择图片")
            return
        for sid in sels:
            if sid in self.image_data:
                del self.image_data[sid]
                self.tree.delete(sid)
        self.log(f"已删除 {len(sels)} 项")
        self.update_stats()
        
    def clear_all(self):
        """清空列表"""
        if not self.image_data: return
        if messagebox.askyesno("确认", "清空所有列表？"):
            self.image_data.clear()
            self.tree.delete(*self.tree.get_children())
            self.update_stats()
            self.log("列表已清空")
            
    def apply_global_settings(self):
        """应用全局设置"""
        fmt = self.output_format.get()
        ratio = self.global_ratio.get()
        
        for item_id, data in self.image_data.items():
            data["ratio"].set(ratio)
            data["format"].set(fmt)
            self.update_tree_row(item_id)
            
        self.log(f"已应用全局设置: {fmt}, 质量 {ratio}%")
        
    def apply_custom_settings(self):
        """应用自定义设置"""
        sels = self.tree.selection()
        if not sels:
            messagebox.showwarning("提示", "请先选择图片")
            return
            
        fmt = self.custom_format.get()
        ratio = self.custom_ratio.get()
        
        for sid in sels:
            if sid in self.image_data:
                self.image_data[sid]["ratio"].set(ratio)
                self.image_data[sid]["format"].set(fmt)
                self.update_tree_row(sid)
                
        self.log(f"已应用自定义设置到 {len(sels)} 张图片")
        
    def update_tree_row(self, item_id):
        """更新Treeview行显示"""
        if item_id not in self.image_data: return
        data = self.image_data[item_id]
        fmt_str = data["format"].get().split(" ")[0]
        self.tree.set(item_id, "输出格式", fmt_str)
        self.tree.set(item_id, "压缩比率", f"{data['ratio'].get()}%")
        
    def update_stats(self):
        """更新统计"""
        total = len(self.image_data)
        size = sum(d["size"] for d in self.image_data.values())
        self.stats_var.set(f"共 {total} 张图片，总大小: {self.fmt_size(size)}")
        
    def start_compress(self):
        """启动压缩线程"""
        if not self.image_data:
            messagebox.showwarning("提示", "请先添加图片")
            return
            
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"创建输出目录失败: {e}")
            return
            
        self.compress_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        
        threading.Thread(target=self.run_compress, daemon=True).start()
        
    def run_compress(self):
        """执行压缩逻辑"""
        total = len(self.image_data)
        success = 0
        fail = 0
        
        self.log(f"开始处理 {total} 张图片...")
        
        for idx, (item_id, data) in enumerate(self.image_data.items(), 1):
            # 更新进度
            self.root.after(0, lambda i=idx, t=total: self.update_progress(i, t))
            
            try:
                result = self.compress_one(item_id, data)
                if result["ok"]:
                    success += 1
                    self.log(f"成功: {data['name']} -> {self.fmt_size(result['size'])} (节省 {result['saved']:.1f}%)")
                    # 更新UI
                    self.root.after(0, lambda iid=item_id, res=result: self.update_row_success(iid, res))
                else:
                    fail += 1
                    self.log(f"失败: {data['name']} - {result['msg']}", "ERROR")
                    self.root.after(0, lambda iid=item_id: self.tree.set(iid, "状态", "失败"))
            except Exception as e:
                fail += 1
                self.log(f"异常: {data['name']} - {e}", "ERROR")
                self.root.after(0, lambda iid=item_id: self.tree.set(iid, "状态", "异常"))
                
        summary = f"处理完成: 成功 {success}, 失败 {fail}"
        self.root.after(0, lambda: self.compress_done(summary))
        
    def compress_one(self, item_id, data):
        """处理单张图片压缩逻辑"""
        try:
            src_path = Path(data["path"])
            img = Image.open(src_path)
            
            # 决定输出格式
            target_fmt_setting = data["format"].get()
            if "JPEG" in target_fmt_setting:
                ext = ".jpg"
                fmt_key = "JPEG"
            elif "WebP" in target_fmt_setting:
                ext = ".webp"
                fmt_key = "WEBP"
            elif "PNG" in target_fmt_setting:
                ext = ".png"
                fmt_key = "PNG"
            elif "ICO" in target_fmt_setting:
                ext = ".ico"
                fmt_key = "ICO"
            elif "BMP" in target_fmt_setting:
                ext = ".bmp"
                fmt_key = "BMP"
            else:
                # 保持原格式
                ext = src_path.suffix.lower()
                if ext in [".jpg", ".jpeg"]:
                    fmt_key = "JPEG"
                elif ext == ".webp":
                    fmt_key = "WEBP"
                elif ext == ".png":
                    fmt_key = "PNG"
                elif ext == ".ico":
                    fmt_key = "ICO"
                elif ext == ".bmp":
                    fmt_key = "BMP"
                else:
                    fmt_key = "PNG" # 默认兜底
            
            quality = data["ratio"].get()
            
            # 构造输出文件名
            out_name = src_path.stem + ext
            out_path = self.output_dir / out_name
            counter = 1
            while out_path.exists():
                out_path = self.output_dir / f"{src_path.stem}_{counter}{ext}"
                counter += 1
            
            # 处理转换逻辑
            save_opts = {}
            
            # 通道转换预处理
            if fmt_key in ["JPEG", "BMP"]:
                # JPEG/BMP 不支持透明，必须转RGB并填充背景
                if img.mode in ("RGBA", "P", "LA"):
                    bg = Image.new("RGB", img.size, (255, 255, 255))
                    if img.mode == "P":
                        img = img.convert("RGBA")
                    bg.paste(img, mask=img.split()[-1]) # 使用Alpha通道作为mask
                    img = bg
                elif img.mode != "RGB":
                    img = img.convert("RGB")
            
            # ICO 特殊处理：限制最大尺寸以保证兼容性
            if fmt_key == "ICO":
                # Windows 标准ICO最大256x256
                if max(img.size) > 256:
                    img.thumbnail((256, 256), Image.Resampling.LANCZOS)
                # ICO格式保存参数
                save_opts['format'] = 'ICO'
                save_opts['sizes'] = [(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)]
            
            # 设置保存参数
            if fmt_key == "JPEG":
                save_opts['format'] = 'JPEG'
                save_opts['quality'] = quality
                save_opts['optimize'] = True
            elif fmt_key == "WEBP":
                save_opts['format'] = 'WEBP'
                save_opts['quality'] = quality
            elif fmt_key == "PNG":
                save_opts['format'] = 'PNG'
                # PNG质量参数(0-9), 值越小压缩越少(文件大但快), 值越大压缩越多(文件小但慢)
                # 这里的质量百分比反向映射到 compress_level
                # 用户选quality=90%, 对应 compress_level=1 (低压缩)
                # 用户选quality=10%, 对应 compress_level=9 (高压缩)
                level = int(9 - (quality / 100) * 9)
                save_opts['compress_level'] = level
            elif fmt_key == "BMP":
                save_opts['format'] = 'BMP'
            
            # 保存
            img.save(out_path, **save_opts)
            
            out_size = out_path.stat().st_size
            saved_percent = (1 - out_size / data["size"]) * 100 if data["size"] > 0 else 0
            
            return {
                "ok": True, 
                "size": out_size, 
                "saved": saved_percent,
                "ratio": quality
            }
            
        except Exception as e:
            return {"ok": False, "msg": str(e)}
            
    def update_row_success(self, item_id, res):
        """更新成功行"""
        self.tree.set(item_id, "压缩后大小", self.fmt_size(res["size"]))
        self.tree.set(item_id, "压缩率", f"{res['saved']:.1f}%")
        self.tree.set(item_id, "状态", "成功")
        # 闪烁效果或颜色标记可选
        
    def update_progress(self, current, total):
        """更新进度条"""
        val = (current / total) * 100
        self.progress_var.set(val)
        self.progress_label.config(text=f"{current}/{total}")
        
    def compress_done(self, msg):
        """压缩完成回调"""
        self.compress_btn.config(state=tk.NORMAL)
        self.progress_var.set(100)
        self.log(msg)
        if messagebox.askyesno("完成", f"{msg}\n\n是否打开输出目录？"):
            self.open_output_dir()
            
    def open_output_dir(self):
        """打开输出目录"""
        try:
            self.output_dir.mkdir(exist_ok=True)
            import platform
            if platform.system() == "Windows":
                os.startfile(self.output_dir)
            elif platform.system() == "Darwin":
                os.system(f'open "{self.output_dir}"')
            else:
                os.system(f'xdg-open "{self.output_dir}"')
        except Exception as e:
            messagebox.showerror("错误", f"无法打开目录: {e}")


def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        print("提示: 安装 tkinterdnd2 可启用拖拽功能")
        
    style = ttk.Style()
    try:
        style.theme_use('clam')
    except:
        pass
        
    app = ImageCompressorApp(root)
    
    # 居中
    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")
    
    root.mainloop()


if __name__ == "__main__":
    main()
