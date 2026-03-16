# 依赖安装: pip install Pillow tkinterdnd2

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
from datetime import datetime

# ================= 依赖检测与导入 =================
try:
    from PIL import Image
except ImportError:
    err_root = tk.Tk()
    err_root.withdraw()
    messagebox.showerror(
        "缺少运行依赖", 
        "运行此程序需要安装 Pillow 库！\n\n请打开 CMD 命令行或终端，输入以下命令安装：\npip install Pillow"
    )
    sys.exit(1)

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False
# =================================================

class ImageCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片处理大师 v3.0 - 压缩/转换/缩放")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # 全局字体设置 (让界面在 Windows 下更好看)
        self.default_font = ("Microsoft YaHei", 9) if sys.platform == "win32" else ("sans-serif", 10)
        self.root.option_add("*Font", self.default_font)

        # 数据存储
        self.image_data = {}
        self.item_counter = 0
        
        # 路径设置
        self.script_dir = Path(__file__).parent if "__file__" in dir() else Path.cwd()
        self.output_dir_var = tk.StringVar(value=str(self.script_dir / "压缩输出"))
        
        # 支持的输入格式
        self.supported_formats = {'.jpg', '.jpeg', '.png', '.bmp', '.webp', '.tiff', '.gif', '.ico'}
        
        # 构建界面
        self.setup_ui()
        self.setup_styles()
        
        # 绑定拖拽
        if HAS_DND:
            self.setup_dnd()
            
    def setup_styles(self):
        """配置更美观的 ttk 样式"""
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
            
        style.configure("TButton", padding=5, font=self.default_font)
        style.configure("TLabelframe", font=(self.default_font[0], 10, "bold"))
        style.configure("TLabelframe.Label", foreground="#333333")
        
        # Treeview 样式
        style.configure("Treeview", rowheight=28, font=self.default_font)
        style.configure("Treeview.Heading", font=(self.default_font[0], 10, "bold"), background="#e1e1e1")
        
        # 定义 Treeview 的标签颜色
        self.tree.tag_configure('odd', background='#F8F9FA')
        self.tree.tag_configure('even', background='#FFFFFF')
        self.tree.tag_configure('success', foreground='#28A745', background='#E8F5E9')
        self.tree.tag_configure('error', foreground='#DC3545', background='#FFEBEE')
        
    def setup_ui(self):
        """构建用户界面"""
        # 主布局容器
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # ================= 左侧：文件列表区域 =================
        left_frame = ttk.Frame(main_paned)
        main_paned.add(left_frame, weight=3)
        
        # 工具栏
        toolbar = ttk.Frame(left_frame)
        toolbar.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(toolbar, text="📁 添加文件", command=self.select_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="📂 添加文件夹", command=self.select_folder).pack(side=tk.LEFT, padx=5)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(toolbar, text="🗑 移除选中", command=self.delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="🧹 清空全部", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        
        # 列表 Treeview
        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        cols = ("文件名", "原始大小", "输出格式", "画质", "缩放比例", "处理后大小", "压缩率", "状态")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended")
        
        # 定义列宽和对齐
        col_widths = [220, 80, 80, 60, 80, 90, 70, 80]
        for col, width in zip(cols, col_widths):
            self.tree.heading(col, text=col)
            anchor = tk.W if col == "文件名" else tk.CENTER
            self.tree.column(col, width=width, anchor=anchor)
        
        # 滚动条
        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        
        # 底部：输出目录选择
        dir_frame = ttk.Frame(left_frame)
        dir_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(dir_frame, text="输出目录:").pack(side=tk.LEFT)
        ttk.Entry(dir_frame, textvariable=self.output_dir_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(dir_frame, text="更改目录", command=self.change_output_dir).pack(side=tk.LEFT)
            
        # ================= 右侧：设置面板 =================
        right_frame = ttk.Frame(main_paned, width=320)
        main_paned.add(right_frame, weight=1)
        
        # 1. 输出格式设置
        fmt_frame = ttk.LabelFrame(right_frame, text="输出格式", padding=10)
        fmt_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.output_format = tk.StringVar(value="保持原格式")
        format_combo = ttk.Combobox(fmt_frame, textvariable=self.output_format, state="readonly")
        format_combo['values'] = ("保持原格式", "JPEG (.jpg)", "WebP (.webp)", "PNG (.png)", "ICO (.ico)")
        format_combo.pack(fill=tk.X)
        
        # 2. 压缩与尺寸参数
        param_frame = ttk.LabelFrame(right_frame, text="压缩参数 (全局)", padding=10)
        param_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 画质
        ttk.Label(param_frame, text="导出画质 (1-100%):", foreground="gray").pack(anchor=tk.W)
        q_frame = ttk.Frame(param_frame)
        q_frame.pack(fill=tk.X, pady=(0, 10))
        self.global_quality = tk.IntVar(value=80)
        ttk.Scale(q_frame, from_=1, to=100, variable=self.global_quality, command=lambda _: self.q_label.config(text=f"{self.global_quality.get()}%")).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.q_label = ttk.Label(q_frame, text="80%", width=4)
        self.q_label.pack(side=tk.RIGHT, padx=(5, 0))
        
        # 尺寸缩放
        ttk.Label(param_frame, text="尺寸缩放 (分辨率 1-100%):", foreground="gray").pack(anchor=tk.W)
        s_frame = ttk.Frame(param_frame)
        s_frame.pack(fill=tk.X, pady=(0, 10))
        self.global_scale_pct = tk.IntVar(value=100)
        ttk.Scale(s_frame, from_=1, to=100, variable=self.global_scale_pct, command=lambda _: self.s_label.config(text=f"{self.global_scale_pct.get()}%")).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.s_label = ttk.Label(s_frame, text="100%", width=4)
        self.s_label.pack(side=tk.RIGHT, padx=(5, 0))

        ttk.Button(param_frame, text="⬇ 应用到所有图片", command=self.apply_global_settings).pack(fill=tk.X)
        
        # 3. 自定义设置 (单张)
        custom_frame = ttk.LabelFrame(right_frame, text="独立设置 (选中项)", padding=10)
        custom_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.custom_quality = tk.IntVar(value=80)
        self.custom_scale_pct = tk.IntVar(value=100)
        
        c_q_frame = ttk.Frame(custom_frame)
        c_q_frame.pack(fill=tk.X, pady=2)
        ttk.Label(c_q_frame, text="画质:").pack(side=tk.LEFT)
        ttk.Scale(c_q_frame, from_=1, to=100, variable=self.custom_quality, command=lambda _: self.c_q_label.config(text=f"{self.custom_quality.get()}%")).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.c_q_label = ttk.Label(c_q_frame, text="80%", width=4)
        self.c_q_label.pack(side=tk.RIGHT)

        c_s_frame = ttk.Frame(custom_frame)
        c_s_frame.pack(fill=tk.X, pady=2)
        ttk.Label(c_s_frame, text="缩放:").pack(side=tk.LEFT)
        ttk.Scale(c_s_frame, from_=1, to=100, variable=self.custom_scale_pct, command=lambda _: self.c_s_label.config(text=f"{self.custom_scale_pct.get()}%")).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.c_s_label = ttk.Label(c_s_frame, text="100%", width=4)
        self.c_s_label.pack(side=tk.RIGHT)
        
        ttk.Button(custom_frame, text="✔ 应用到选中图片", command=self.apply_custom_settings).pack(fill=tk.X, pady=(5, 0))

        # 4. 操作区
        action_frame = ttk.LabelFrame(right_frame, text="执行", padding=10)
        action_frame.pack(fill=tk.X, expand=True, anchor=tk.N)
        
        self.compress_btn = tk.Button(action_frame, text="🚀 开始处理", font=("Microsoft YaHei", 12, "bold"), bg="#0078D7", fg="white", relief=tk.FLAT, cursor="hand2", command=self.start_compress)
        self.compress_btn.pack(fill=tk.X, pady=5, ipady=5)
        
        ttk.Button(action_frame, text="📂 打开输出目录", command=self.open_output_dir).pack(fill=tk.X, pady=5)
        
        self.stats_var = tk.StringVar(value="等待添加文件...")
        ttk.Label(action_frame, textvariable=self.stats_var, foreground="gray").pack(anchor=tk.CENTER, pady=(10, 0))

        # ================= 底部：日志和进度 =================
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(bottom_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        self.log_text = tk.Text(bottom_frame, height=4, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 9), bg="#F1F1F1", relief=tk.FLAT)
        self.log_text.pack(fill=tk.X)

    def setup_dnd(self):
        """设置拖拽功能"""
        if not HAS_DND: return
        self.tree.drop_target_register(DND_FILES)
        
        def on_drop(event):
            # 使用 splitlist 完美解析包含空格的文件路径
            paths = self.tree.tk.splitlist(event.data)
            self.add_images_from_paths(paths)
            self.tree.configure(background="white")
            
        self.tree.dnd_bind('<<Drop>>', on_drop)
        self.tree.dnd_bind('<<DragEnter>>', lambda e: self.tree.configure(background="#E3F2FD"))
        self.tree.dnd_bind('<<DragLeave>>', lambda e: self.tree.configure(background="white"))

    def log(self, msg, level="INFO"):
        """写入日志"""
        ts = datetime.now().strftime("%H:%M:%S")
        color = "red" if level == "ERROR" else "black"
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        
    def fmt_size(self, b):
        """格式化文件大小"""
        if b < 1024: return f"{b} B"
        if b < 1024**2: return f"{b/1024:.1f} KB"
        return f"{b/1024**2:.2f} MB"
        
    def change_output_dir(self):
        d = filedialog.askdirectory(title="选择输出目录", initialdir=self.output_dir_var.get())
        if d: self.output_dir_var.set(d)

    def on_tree_select(self, event):
        """列表选择事件，同步右侧设置"""
        sels = self.tree.selection()
        if len(sels) == 1:
            item_id = sels[0]
            if item_id in self.image_data:
                data = self.image_data[item_id]
                self.custom_quality.set(data["quality"].get())
                self.custom_scale_pct.set(data["scale"].get())
                self.c_q_label.config(text=f"{data['quality'].get()}%")
                self.c_s_label.config(text=f"{data['scale'].get()}%")
                
    def select_files(self):
        files = filedialog.askopenfilenames(title="选择图片", filetypes=[("图片文件", "*.jpg *.png *.jpeg *.bmp *.webp *.gif *.ico")])
        if files: self.add_images_from_paths(files)
        
    def select_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder: self.add_images_from_paths([folder])
        
    def add_images_from_paths(self, paths):
        count = 0
        for p in paths:
            path = Path(p)
            if not path.exists(): continue
            if path.is_file():
                if self.add_single_image(path): count += 1
            elif path.is_dir():
                for root, _, files in os.walk(path):
                    for f in files:
                        if self.add_single_image(Path(root) / f): count += 1
        
        self.refresh_list_colors()
        self.update_stats()
        if count > 0:
            self.log(f"成功导入 {count} 张图片")
            
    def add_single_image(self, path):
        if path.suffix.lower() not in self.supported_formats:
            return False
        
        path_str = str(path.resolve())
        # 去重
        if any(d["path"] == path_str for d in self.image_data.values()): 
            return False
            
        size = path.stat().st_size
        item_id = f"img_{self.item_counter}"
        self.item_counter += 1
        
        # 初始参数绑定全局
        q_var = tk.IntVar(value=self.global_quality.get())
        s_var = tk.IntVar(value=self.global_scale_pct.get())
        fmt_var = tk.StringVar(value=self.output_format.get())
        
        self.tree.insert("", tk.END, iid=item_id, values=(
            path.name, self.fmt_size(size), 
            fmt_var.get().split(" ")[0], 
            f"{q_var.get()}%", f"{s_var.get()}%", "-", "-", "待处理"
        ))
        
        self.image_data[item_id] = {
            "path": path_str, "name": path.name, "size": size, 
            "quality": q_var, "scale": s_var, "format": fmt_var
        }
        return True
        
    def refresh_list_colors(self):
        """刷新列表奇偶行颜色"""
        for index, item in enumerate(self.tree.get_children()):
            # 保留原本的状态颜色
            current_tags = self.tree.item(item, "tags")
            if "success" not in current_tags and "error" not in current_tags:
                tag = 'even' if index % 2 == 0 else 'odd'
                self.tree.item(item, tags=(tag,))
                
    def delete_selected(self):
        sels = self.tree.selection()
        if not sels: return
        for sid in sels:
            if sid in self.image_data:
                del self.image_data[sid]
                self.tree.delete(sid)
        self.refresh_list_colors()
        self.update_stats()
        
    def clear_all(self):
        if not self.image_data: return
        self.image_data.clear()
        self.tree.delete(*self.tree.get_children())
        self.update_stats()
        self.log("列表已清空")
            
    def apply_global_settings(self):
        fmt = self.output_format.get()
        q = self.global_quality.get()
        s = self.global_scale_pct.get()
        for item_id, data in self.image_data.items():
            data["quality"].set(q)
            data["scale"].set(s)
            data["format"].set(fmt)
            self.update_tree_row(item_id)
        self.log(f"已应用全局参数: {fmt.split(' ')[0]}, 画质 {q}%, 缩放 {s}%")
        
    def apply_custom_settings(self):
        sels = self.tree.selection()
        if not sels: return
        q = self.custom_quality.get()
        s = self.custom_scale_pct.get()
        for sid in sels:
            if sid in self.image_data:
                self.image_data[sid]["quality"].set(q)
                self.image_data[sid]["scale"].set(s)
                self.update_tree_row(sid)
        self.log(f"已为 {len(sels)} 张图片应用独立参数")
        
    def update_tree_row(self, item_id):
        data = self.image_data[item_id]
        self.tree.set(item_id, "输出格式", data["format"].get().split(" ")[0])
        self.tree.set(item_id, "画质", f"{data['quality'].get()}%")
        self.tree.set(item_id, "缩放比例", f"{data['scale'].get()}%")
        
    def update_stats(self):
        total = len(self.image_data)
        size = sum(d["size"] for d in self.image_data.values())
        self.stats_var.set(f"共 {total} 个文件 | 预估总体积: {self.fmt_size(size)}")
        
    def start_compress(self):
        if not self.image_data:
            messagebox.showwarning("提示", "列表为空，请先添加图片！")
            return
            
        out_path_str = self.output_dir_var.get()
        if not out_path_str.strip():
            messagebox.showerror("错误", "输出目录不能为空！")
            return
            
        out_dir = Path(out_path_str)
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"创建输出目录失败:\n{e}")
            return
            
        self.compress_btn.config(state=tk.DISABLED, text="处理中...", bg="gray")
        self.progress_var.set(0)
        self.log("========== 开始批量处理 ==========")
        
        # 重置所有状态颜色
        self.refresh_list_colors()
        
        threading.Thread(target=self.run_compress, args=(out_dir,), daemon=True).start()
        
    def run_compress(self, out_dir):
        total = len(self.image_data)
        success = 0
        fail = 0
        
        for idx, (item_id, data) in enumerate(self.image_data.items(), 1):
            # 更新进度条
            self.root.after(0, lambda v=idx/total*100: self.progress_var.set(v))
            
            try:
                result = self.compress_one(data, out_dir)
                if result["ok"]:
                    success += 1
                    self.root.after(0, self.update_row_success, item_id, result)
                    self.log(f"✔ {data['name']} -> {self.fmt_size(result['size'])}")
                else:
                    fail += 1
                    self.root.after(0, self.update_row_fail, item_id, result['msg'])
                    self.log(f"✖ {data['name']} 处理失败: {result['msg']}", "ERROR")
            except Exception as e:
                fail += 1
                self.root.after(0, self.update_row_fail, item_id, str(e))
                self.log(f"✖ {data['name']} 严重异常: {e}", "ERROR")
                
        summary = f"处理完成! 成功: {success} 张, 失败: {fail} 张。"
        self.root.after(0, self.compress_done, summary)
        
    def compress_one(self, data, out_dir: Path):
        try:
            src_path = Path(data["path"])
            if not src_path.exists():
                return {"ok": False, "msg": "源文件已丢失"}
                
            img = Image.open(src_path)
            
            # 解析格式
            target_fmt = data["format"].get()
            ext = src_path.suffix.lower()
            fmt_key = "JPEG" if ext in ['.jpg', '.jpeg'] else ext[1:].upper()
            
            if "JPEG" in target_fmt: ext, fmt_key = ".jpg", "JPEG"
            elif "WebP" in target_fmt: ext, fmt_key = ".webp", "WEBP"
            elif "PNG" in target_fmt: ext, fmt_key = ".png", "PNG"
            elif "ICO" in target_fmt: ext, fmt_key = ".ico", "ICO"
            
            # 1. 处理缩放 (Resize)
            scale = data["scale"].get()
            if scale < 100:
                new_w = max(1, int(img.width * (scale / 100)))
                new_h = max(1, int(img.height * (scale / 100)))
                img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            
            # 2. 处理通道 (透明度兼容)
            if fmt_key == "JPEG":
                if img.mode in ("RGBA", "P", "LA"):
                    bg = Image.new("RGB", img.size, (255, 255, 255))
                    if img.mode == "P": img = img.convert("RGBA")
                    if img.mode == "RGBA" or img.mode == "LA":
                        bg.paste(img, mask=img.split()[-1])
                    img = bg
                elif img.mode != "RGB":
                    img = img.convert("RGB")
                    
            # 3. 输出路径与重命名
            out_name = src_path.stem + ext
            out_path = out_dir / out_name
            counter = 1
            while out_path.exists():
                out_path = out_dir / f"{src_path.stem}_{counter}{ext}"
                counter += 1
                
            # 4. 保存参数构建
            save_opts = {'format': fmt_key}
            quality = data["quality"].get()
            
            if fmt_key in ["JPEG", "WEBP"]:
                save_opts['quality'] = quality
                if fmt_key == "JPEG": save_opts['optimize'] = True
            elif fmt_key == "PNG":
                # PNG optimize 减小体积，不损失画质但处理较慢
                save_opts['optimize'] = True 
            elif fmt_key == "ICO":
                if max(img.size) > 256:
                    img.thumbnail((256, 256), Image.Resampling.LANCZOS)
                save_opts['sizes'] = [(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)]
                
            img.save(out_path, **save_opts)
            
            out_size = out_path.stat().st_size
            saved_pct = (1 - out_size / data["size"]) * 100 if data["size"] > 0 else 0
            
            return {"ok": True, "size": out_size, "saved": saved_pct}
            
        except Exception as e:
            return {"ok": False, "msg": str(e)}
            
    def update_row_success(self, item_id, res):
        self.tree.set(item_id, "处理后大小", self.fmt_size(res["size"]))
        # 根据压缩率显示不同符号
        sign = "↓" if res['saved'] >= 0 else "↑"
        self.tree.set(item_id, "压缩率", f"{sign} {abs(res['saved']):.1f}%")
        self.tree.set(item_id, "状态", "✅ 成功")
        self.tree.item(item_id, tags=("success",))
        
    def update_row_fail(self, item_id, msg):
        self.tree.set(item_id, "状态", "❌ 失败")
        self.tree.item(item_id, tags=("error",))
        
    def compress_done(self, msg):
        self.compress_btn.config(state=tk.NORMAL, text="🚀 开始处理", bg="#0078D7")
        self.progress_var.set(100)
        self.log("========== 处理结束 ==========")
        if messagebox.askyesno("处理完成", f"{msg}\n\n是否打开输出目录？"):
            self.open_output_dir()
            
    def open_output_dir(self):
        try:
            d = self.output_dir_var.get()
            Path(d).mkdir(parents=True, exist_ok=True)
            if sys.platform == "win32":
                os.startfile(d)
            elif sys.platform == "darwin":
                os.system(f'open "{d}"')
            else:
                os.system(f'xdg-open "{d}"')
        except Exception as e:
            messagebox.showerror("错误", f"无法打开目录: {e}")

def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        print("提示: 未安装 tkinterdnd2，无法启用拖拽功能")
        
    app = ImageCompressorApp(root)
    
    # 屏幕居中
    root.update_idletasks()
    w = root.winfo_width()
    h = root.winfo_height()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")
    
    root.mainloop()

if __name__ == "__main__":
    main()
