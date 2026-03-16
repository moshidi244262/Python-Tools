import sys
import os
import shutil
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QLabel, QTextEdit, QFileDialog, QProgressBar,
                               QFrame, QMessageBox)
from PySide6.QtCore import Qt, QThread, Signal, QUrl
from PySide6.QtGui import QFont, QIcon, QDesktopServices
import mutagen

# ================= 样式表 (美化 UI) =================
STYLESHEET = """
QWidget {
    font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
    color: #333333;
    background-color: #F5F7FA;
}

QFrame#ControlPanel {
    background-color: #FFFFFF;
    border-radius: 10px;
    border: 1px solid #E4E7ED;
}

QPushButton {
    background-color: #409EFF;
    color: white;
    border: none;
    border-radius: 6px;
    padding: 8px 16px;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #66B1FF;
}

QPushButton:pressed {
    background-color: #3A8EE6;
}

QPushButton:disabled {
    background-color: #A0CFFF;
    color: #FFFFFF;
}

QPushButton#StartBtn {
    background-color: #67C23A;
}
QPushButton#StartBtn:hover {
    background-color: #85CE61;
}
QPushButton#StartBtn:pressed {
    background-color: #5DAF34;
}
QPushButton#StartBtn:disabled {
    background-color: #B3E19D;
}

QPushButton#StopBtn {
    background-color: #F56C6C;
}
QPushButton#StopBtn:hover {
    background-color: #F89898;
}
QPushButton#StopBtn:pressed {
    background-color: #E6A23C;
}

QPushButton#OpenDirBtn {
    background-color: #E6A23C;
}
QPushButton#OpenDirBtn:hover {
    background-color: #EBB563;
}
QPushButton#OpenDirBtn:pressed {
    background-color: #CF9236;
}

QProgressBar {
    border: 1px solid #E4E7ED;
    border-radius: 6px;
    text-align: center;
    background-color: #EBEEF5;
    color: #303133;
    font-weight: bold;
}

QProgressBar::chunk {
    background-color: #67C23A;
    border-radius: 5px;
}

QTextEdit {
    background-color: #282C34;
    color: #ABB2BF;
    border-radius: 8px;
    padding: 10px;
    font-family: "Consolas", "Courier New", monospace;
    font-size: 13px;
    border: 1px solid #DCDFE6;
}

QLabel#DropZone {
    border: 2px dashed #C0C4CC;
    border-radius: 12px;
    background-color: #FAFAFA;
    padding: 30px;
    color: #909399;
}
"""


class TagCleanerWorker(QThread):
    """
    后台处理线程
    """
    log_signal = Signal(str)
    progress_signal = Signal(int, int)
    finished_signal = Signal(dict) # 返回统计数据

    def __init__(self, paths):
        super().__init__()
        self.paths = paths
        # 扩展了支持的格式
        self.target_exts = {'.mp3', '.wav', '.flac', '.m4a', '.ogg', '.wma', '.aac'}
        self.is_cancelled = False # 用于控制停止

    def cancel(self):
        self.is_cancelled = True

    def run(self):
        stats = {'success': 0, 'skipped': 0, 'failed': 0}
        
        # 获取当前脚本所在目录并创建“去标签音乐”文件夹
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        output_dir = os.path.join(script_dir, "去标签音乐")
        
        try:
            os.makedirs(output_dir, exist_ok=True)
            self.log_signal.emit(f"📁 目标输出文件夹: {output_dir}\n")
        except Exception as e:
            self.log_signal.emit(f"❌ 创建输出目录失败: {str(e)}")
            self.finished_signal.emit(stats)
            return

        # 1. 收集所有需要处理的文件
        files_to_process = []
        for path in self.paths:
            if os.path.isfile(path):
                if os.path.splitext(path)[1].lower() in self.target_exts:
                    files_to_process.append(path)
            elif os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for f in files:
                        if os.path.splitext(f)[1].lower() in self.target_exts:
                            files_to_process.append(os.path.join(root, f))

        total_files = len(files_to_process)
        if total_files == 0:
            self.log_signal.emit("⚠️ 没有找到支持的音频文件。")
            self.finished_signal.emit(stats)
            return

        self.log_signal.emit(f"🔍 共确认 {total_files} 个音频文件，开始处理...\n{'-'*40}")

        # 2. 逐个文件复制并清除标签
        for i, file_path in enumerate(files_to_process):
            if self.is_cancelled:
                self.log_signal.emit("\n🛑 用户手动终止了操作！")
                break

            filename = os.path.basename(file_path)
            
            # 处理文件名冲突 (如果输出目录已经存在同名文件，添加后缀)
            out_path = os.path.join(output_dir, filename)
            counter = 1
            while os.path.exists(out_path):
                name, ext = os.path.splitext(filename)
                out_path = os.path.join(output_dir, f"{name}_{counter}{ext}")
                counter += 1
                
            out_filename = os.path.basename(out_path)

            try:
                # 复制原文件到目标文件夹，保留原文件不动
                shutil.copy2(file_path, out_path)
                
                # 对复制后的新文件进行标签清除
                audio = mutagen.File(out_path)
                
                if audio is None:
                    self.log_signal.emit(f"[跳过] {filename} -> 无法解析格式")
                    os.remove(out_path) # 清理无法处理的复制文件
                    stats['skipped'] += 1
                elif audio.tags is None or len(audio.tags) == 0:
                    self.log_signal.emit(f"[跳过] {filename} -> 已经是纯净版 (已复制为 {out_filename})")
                    stats['skipped'] += 1
                else:
                    audio.delete()
                    if hasattr(audio, 'save'):
                        audio.save() # 某些格式需要显式保存
                    self.log_signal.emit(f"[成功] {filename} -> 标签已清空 (已保存为 {out_filename})")
                    stats['success'] += 1
                    
            except Exception as e:
                self.log_signal.emit(f"[失败] {filename} -> 错误: {str(e)}")
                if os.path.exists(out_path):
                    os.remove(out_path) # 清理失败的遗留文件
                stats['failed'] += 1

            self.progress_signal.emit(i + 1, total_files)

        self.finished_signal.emit(stats)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🎵 音乐标签极净清除工具 v3.1")
        self.resize(800, 600)
        self.setAcceptDrops(True)
        self.setStyleSheet(STYLESHEET)
        
        self.worker = None
        self.is_processing = False
        self.queued_paths = set() # 待处理文件队列
        
        self.setup_ui()

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 1. 拖拽提示区域
        self.drop_label = QLabel("将音乐文件或文件夹拖拽到此处加入队列\n\n支持格式: MP3, WAV, FLAC, M4A, OGG 等\n(处理后会自动保存在同目录下的“去标签音乐”文件夹)")
        self.drop_label.setObjectName("DropZone")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setFont(QFont("Microsoft YaHei", 12, QFont.Bold))
        self.drop_label.setMinimumHeight(150)
        main_layout.addWidget(self.drop_label)

        # 2. 控制面板 (包裹在 QFrame 中产生卡片效果)
        control_panel = QFrame()
        control_panel.setObjectName("ControlPanel")
        control_layout = QVBoxLayout(control_panel)
        control_layout.setContentsMargins(15, 15, 15, 15)
        control_layout.setSpacing(10)

        # 按钮行
        btn_layout = QHBoxLayout()
        self.btn_select_files = QPushButton("📄 添加文件")
        self.btn_select_folder = QPushButton("📁 添加文件夹")
        
        self.btn_start = QPushButton("▶️ 开始去除")
        self.btn_start.setObjectName("StartBtn")
        self.btn_start.setEnabled(False) # 初始禁用，直到有文件加入队列
        
        self.btn_stop = QPushButton("🛑 停止")
        self.btn_stop.setObjectName("StopBtn")
        self.btn_stop.setVisible(False) # 初始隐藏停止按钮

        # 新增：打开输出目录按钮
        self.btn_open_dir = QPushButton("📂 打开输出目录")
        self.btn_open_dir.setObjectName("OpenDirBtn")
        
        self.btn_clear_log = QPushButton("🗑️ 清空日志")
        
        self.btn_select_files.clicked.connect(self.select_files)
        self.btn_select_folder.clicked.connect(self.select_folder)
        self.btn_start.clicked.connect(self.start_processing)
        self.btn_stop.clicked.connect(self.stop_processing)
        self.btn_open_dir.clicked.connect(self.open_output_dir)
        self.btn_clear_log.clicked.connect(lambda: self.log_area.clear())

        btn_layout.addWidget(self.btn_select_files)
        btn_layout.addWidget(self.btn_select_folder)
        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(self.btn_stop)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_open_dir)
        btn_layout.addWidget(self.btn_clear_log)
        
        control_layout.addLayout(btn_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setMinimumHeight(25)
        self.progress_bar.setFormat("等待添加文件...")
        control_layout.addWidget(self.progress_bar)

        main_layout.addWidget(control_panel)

        # 3. 日志输出区域
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        main_layout.addWidget(self.log_area)

        self.log_msg("✨ 欢迎使用音乐标签极净清除工具！\n提示：此版本将会保留您的原文件，并在本程序目录下自动创建【去标签音乐】文件夹存放处理后的音乐。\n\n请添加文件或文件夹至队列...\n")

    def log_msg(self, msg):
        self.log_area.append(msg)
        self.log_area.verticalScrollBar().setValue(self.log_area.verticalScrollBar().maximum())

    def update_progress(self, current, total):
        percentage = int((current / total) * 100)
        self.progress_bar.setValue(percentage)
        self.progress_bar.setFormat(f"处理中: {current}/{total} ({percentage}%)")

    def queue_paths(self, paths):
        """将选择的路径加入待处理队列"""
        if not paths or self.is_processing:
            return
            
        added_count = 0
        for p in paths:
            if p not in self.queued_paths:
                self.queued_paths.add(p)
                added_count += 1
                
        if added_count > 0:
            self.log_msg(f"📥 成功添加 {added_count} 个路径，当前队列共 {len(self.queued_paths)} 项。请点击【▶️ 开始去除】执行。")
            self.btn_start.setEnabled(True)
            self.progress_bar.setFormat(f"等待开始... (队列中有 {len(self.queued_paths)} 项)")

    def start_processing(self):
        """开始处理队列中的文件"""
        if not self.queued_paths or self.is_processing:
            return
            
        self.is_processing = True
        self.btn_select_files.setVisible(False)
        self.btn_select_folder.setVisible(False)
        self.btn_start.setVisible(False)
        self.btn_stop.setVisible(True)
        
        self.progress_bar.setValue(0)
        
        # 启动后台线程
        self.worker = TagCleanerWorker(list(self.queued_paths))
        self.worker.log_signal.connect(self.log_msg)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.on_processing_finished)
        self.worker.start()

    def stop_processing(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.btn_stop.setText("正在停止...")
            self.btn_stop.setEnabled(False)

    def open_output_dir(self):
        """使用系统默认文件管理器打开输出目录"""
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        output_dir = os.path.join(script_dir, "去标签音乐")
        # 如果目录尚不存在则先创建
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        # 调用 QDesktopServices 跨平台打开文件夹
        QDesktopServices.openUrl(QUrl.fromLocalFile(output_dir))

    def on_processing_finished(self, stats):
        self.is_processing = False
        self.btn_select_files.setVisible(True)
        self.btn_select_folder.setVisible(True)
        self.btn_stop.setVisible(False)
        
        self.btn_start.setVisible(True)
        self.btn_start.setEnabled(False) # 处理完后队列清空，禁用开始按钮
        
        self.btn_stop.setText("🛑 停止")
        self.btn_stop.setEnabled(True)
        
        # 清空队列
        self.queued_paths.clear()
        
        if stats['success'] > 0 or stats['failed'] > 0 or stats['skipped'] > 0:
            summary = f"\n{'-'*40}\n🎉 处理结束！统计信息：\n"
            summary += f"✅ 成功清除并保存: {stats['success']} 个\n"
            summary += f"⏩ 无需处理并复制: {stats['skipped']} 个\n"
            summary += f"❌ 处理失败/不支持: {stats['failed']} 个\n"
            self.log_msg(summary)
            self.progress_bar.setFormat("处理完成，队列已清空")
        else:
            self.progress_bar.setFormat("无有效文件，队列已清空")

    # --- 按钮事件 ---
    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择音频文件加入队列", "", "Audio Files (*.mp3 *.wav *.flac *.m4a *.ogg *.wma *.aac)"
        )
        if files:
            self.queue_paths(files)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹加入队列")
        if folder:
            self.queue_paths([folder])

    # --- 拖拽事件支持 ---
    def dragEnterEvent(self, event):
        # 优化：如果在处理中，拒绝拖入，防止线程冲突
        if self.is_processing:
            event.ignore()
            return

        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_label.setStyleSheet("""
                QLabel#DropZone {
                    border: 2px solid #67C23A;
                    background-color: #F0F9EB;
                    color: #67C23A;
                }
            """)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.drop_label.setStyleSheet("") # 恢复 QSS 默认样式

    def dropEvent(self, event):
        self.dragLeaveEvent(event)
        if self.is_processing:
            return
            
        urls = event.mimeData().urls()
        # 兼容 Windows 路径
        paths = [url.toLocalFile() for url in urls if url.isLocalFile()]
        if paths:
            self.queue_paths(paths)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 强制启用高 DPI 支持 (让字体在 4K 屏幕上不模糊)
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())
