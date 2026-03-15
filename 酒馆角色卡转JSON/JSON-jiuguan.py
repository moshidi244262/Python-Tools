# 依赖安装: pip install Pillow tkinterdnd2

import os
import json
import base64
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import re
from PIL import Image
from tkinterdnd2 import TkinterDnD, DND_FILES

# 配置常量
OUTPUT_DIR = r"C:\Users\24426\Desktop\Py工具\酒馆角色卡转JSON\角色信息"

class CharacterCardExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("酒馆角色卡信息提取工具 v2.0 (修复版)")
        self.root.geometry("800x600")
        
        # 确保输出目录存在
        try:
            if not os.path.exists(OUTPUT_DIR):
                os.makedirs(OUTPUT_DIR)
                self.log(f"已创建输出目录: {OUTPUT_DIR}")
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出目录: {e}")
            root.destroy()
            return

        # --- UI 布局 ---
        # 顶部按钮区
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Button(btn_frame, text="选择文件 (支持多选)", command=self.select_files, width=20).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="选择文件夹 (递归)", command=self.select_folder, width=20).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="清空日志", command=self.clear_log, width=10).pack(side=tk.RIGHT, padx=5)

        # 拖拽提示区
        # 注意：为了拖拽稳定，我们将整个窗口背景或一个大控件作为拖放区
        self.drop_label = tk.Label(root, text="👇 将角色卡文件或文件夹拖拽到此处 👇", bg="#e1f5fe", font=("Arial", 14), height=2, relief="groove")
        self.drop_label.pack(fill=tk.X, padx=10, pady=5)

        # 日志显示区
        self.log_area = scrolledtext.ScrolledText(root, state='disabled', font=("Consolas", 10))
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 底部状态栏
        self.status_var = tk.StringVar(value="就绪")
        tk.Label(root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W).pack(side=tk.BOTTOM, fill=tk.X)

        # --- 绑定拖拽事件 ---
        # 注册拖拽目标
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)
        
        # 额外绑定：防呆设计，有些系统拖拽到子控件上不响应，绑定到root更稳妥
        # 但tkinterdnd2通常需要显式注册控件，这里主要优化数据处理逻辑

    def log(self, message):
        """线程安全的日志输出"""
        def _log():
            self.log_area.config(state='normal')
            self.log_area.insert(tk.END, message + "\n")
            self.log_area.see(tk.END)
            self.log_area.config(state='disabled')
        self.root.after(0, _log)

    def clear_log(self):
        self.log_area.config(state='normal')
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state='disabled')

    def select_files(self):
        """选择文件对话框"""
        files = filedialog.askopenfilenames(
            title="选择角色卡文件",
            filetypes=[("图片文件", "*.png *.webp"), ("所有文件", "*.*")]
        )
        if files:
            self.start_processing(list(files))

    def select_folder(self):
        """选择文件夹对话框"""
        folder = filedialog.askdirectory(title="选择包含角色卡的文件夹")
        if folder:
            self.start_processing([folder])

    def on_drop(self, event):
        """处理拖拽事件 - 增强路径清洗逻辑"""
        # 获取原始路径数据
        raw_data = event.data
        
        # 修复Bug：Windows下路径可能包含大括号 {} 且被转义
        # tkinterdnd2 在 Windows 上有时返回 {C:/path} 格式
        paths = []
        
        # 这里的 splitlist 已经处理了大部分分割，但在 Windows 路径含空格时可能有坑
        # 我们手动清洗一下
        try:
            # 尝试使用 tk 的内置分割
            split_paths = self.root.tk.splitlist(raw_data)
            for p in split_paths:
                # 去除首尾可能存在的花括号
                p_clean = p.strip()
                if p_clean.startswith('{') and p_clean.endswith('}'):
                    p_clean = p_clean[1:-1]
                paths.append(p_clean)
        except Exception as e:
            self.log(f"拖拽数据解析异常: {e}")
            # 简单的备用分割
            paths = [raw_data.strip('{}')]

        if paths:
            # 打印解析出的路径以便调试
            self.log(f"接收到拖拽: {len(paths)} 个目标")
            self.start_processing(paths)

    def start_processing(self, targets):
        """启动处理线程，防止界面卡死"""
        self.status_var.set("正在处理中...")
        thread = threading.Thread(target=self.process_targets, args=(targets,))
        thread.daemon = True
        thread.start()

    def process_targets(self, targets):
        """处理文件/文件夹列表"""
        all_files = []
        
        # 1. 收集所有文件路径
        for path in targets:
            if not path:
                continue
            
            if os.path.isfile(path):
                all_files.append(path)
            elif os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for file in files:
                        all_files.append(os.path.join(root, file))
        
        # 过滤非图片文件
        valid_files = [f for f in all_files if f.lower().endswith(('.png', '.webp'))]
        
        if not valid_files:
            self.log("未发现 .png 或 .webp 文件。")
            self.status_var.set("就绪")
            return

        self.log(f"发现 {len(valid_files)} 个图片文件待解析...")

        success_count = 0
        fail_count = 0

        # 2. 逐个提取
        for file_path in valid_files:
            try:
                result = self.extract_metadata(file_path)
                if result:
                    # 保存文件
                    base_name = os.path.splitext(os.path.basename(file_path))[0]
                    # 清理文件名中的非法字符
                    safe_name = "".join([c for c in base_name if c.isalnum() or c in (' ', '_', '-', '(', ')')]).strip()
                    if not safe_name: safe_name = "Unnamed_Character"
                    
                    out_file = os.path.join(OUTPUT_DIR, f"{safe_name}.json")
                    
                    # 防止覆盖，添加序号
                    counter = 1
                    while os.path.exists(out_file):
                        out_file = os.path.join(OUTPUT_DIR, f"{safe_name}_{counter}.json")
                        counter += 1
                    
                    with open(out_file, 'w', encoding='utf-8') as f:
                        json.dump(result, f, ensure_ascii=False, indent=4)
                    
                    self.log(f"[成功] {os.path.basename(file_path)} -> {os.path.basename(out_file)}")
                    success_count += 1
                else:
                    self.log(f"[跳过] {os.path.basename(file_path)} (非角色卡或无数据)")
                    fail_count += 1
            except Exception as e:
                self.log(f"[错误] {os.path.basename(file_path)}: {str(e)}")
                fail_count += 1

        self.log("-" * 50)
        self.log(f"处理完成。成功: {success_count}, 失败/跳过: {fail_count}")
        self.status_var.set("就绪")

    def extract_metadata(self, file_path):
        """从图片中提取角色卡数据 - 增强版解析"""
        try:
            with Image.open(file_path) as img:
                # 确保图像已加载（某些懒加载模式可能导致info为空）
                img.load()
                
                # 获取元数据字典
                metadata = img.info
                
                if not metadata:
                    # 如果没有元数据，直接返回
                    return None

                # 尝试查找 'chara' 字段
                # 某些系统的键名可能大小写不同，进行一次遍历查找
                chara_data = None
                for key, value in metadata.items():
                    if key.lower() == 'chara':
                        chara_data = value
                        break
                
                if not chara_data:
                    # 深度解析：如果标准字段不存在，尝试寻找所有可能的 Base64 串
                    # 仅当标准查找失败时启用，防止误判
                    for key, value in metadata.items():
                        if isinstance(value, str) and len(value) > 100:
                            # 尝试解码看看是不是 JSON
                            try:
                                # 尝试 Base64 解码
                                decoded = base64.b64decode(value)
                                # 检查解码后的内容是否像 JSON
                                if b'"name"' in decoded or b'"spec"' in decoded:
                                    chara_data = value
                                    self.log(f"  > 检测到非标准键名: {key}")
                                    break
                            except:
                                continue
                
                if not chara_data:
                    return None

                # 解码 Base64
                # 修正 Padding：Base64 字符串长度必须是 4 的倍数
                # 很多角色卡数据缺少末尾的 '=' 补位，导致解码失败，这是常见的 Bug 来源
                missing_padding = len(chara_data) % 4
                if missing_padding:
                    chara_data += '=' * (4 - missing_padding)
                
                # 尝试解码
                decoded_bytes = base64.b64decode(chara_data)
                
                # 尝试解析 JSON
                # 某些旧卡可能是 UTF-16 或其他编码，但绝大多数是 UTF-8
                text_content = decoded_bytes.decode('utf-8')
                
                json_data = json.loads(text_content)
                
                # 兼容 V2 卡片结构：
                # V2 卡片结构为 {"spec": "chara_card_v2", "data": { ... }}
                # 我们需要提取其中的 data 部分，或者直接返回整个对象
                # 为了信息完整性，直接返回解析出的整个对象
                return json_data

        except json.JSONDecodeError as e:
            self.log(f"  > JSON解析失败: {e}")
            return None
        except base64.binascii.Error as e:
            self.log(f"  > Base64解码失败，数据格式错误: {e}")
            return None
        except Exception as e:
            self.log(f"  > 未知解析错误: {e}")
            return None

if __name__ == "__main__":
    # 检查依赖
    try:
        from PIL import Image
        from tkinterdnd2 import TkinterDnD
    except ImportError as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("缺少依赖", f"运行程序缺少必要的库: {e}\n\n请在终端运行:\npip install Pillow tkinterdnd2")
        exit(1)

    root = TkinterDnD.Tk()
    app = CharacterCardExtractor(root)
    root.mainloop()
