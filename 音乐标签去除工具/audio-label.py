import sys
import os
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QLabel, QTextEdit, QFileDialog, QProgressBar)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont
import mutagen

class TagCleanerWorker(QThread):
    """
    后台处理线程，防止处理大量文件时导致主界面卡顿
    """
    log_signal = Signal(str)         # 发送日志信息的信号
    progress_signal = Signal(int, int) # 发送进度(当前, 总数)的信号
    finished_signal = Signal()       # 处理完成的信号

    def __init__(self, paths):
        super().__init__()
        self.paths = paths
        # 支持的音频格式
        self.target_exts = {'.mp3', '.wav', '.flac'}

    def run(self):
        # 1. 收集所有需要处理的文件
        files_to_process = []
        for path in self.paths:
            if os.path.isfile(path):
                if os.path.splitext(path)[1].lower() in self.target_exts:
                    files_to_process.append(path)
            elif os.path.isdir(path):
                # 递归遍历文件夹及子文件夹
                for root, dirs, files in os.walk(path):
                    for f in files:
                        if os.path.splitext(f)[1].lower() in self.target_exts:
                            files_to_process.append(os.path.join(root, f))

        total_files = len(files_to_process)
        if total_files == 0:
            self.log_signal.emit("⚠️ 没有找到支持的音频文件 (MP3, WAV, FLAC)。")
            self.finished_signal.emit()
            return

        self.log_signal.emit(f"🔍 共找到 {total_files} 个音频文件，开始深度清除标签...")

        # 2. 逐个文件清除标签
        for i, file_path in enumerate(files_to_process):
            filename = os.path.basename(file_path)
            try:
                # 使用 mutagen 读取音频文件
                audio = mutagen.File(file_path)
                if audio is not None:
                    # delete() 方法会彻底清空该文件的所有内嵌标签数据（保留纯音频流）
                    audio.delete() 
                    self.log_signal.emit(f"[成功] {filename} -> 标签已清空")
                else:
                    self.log_signal.emit(f"[跳过] {filename} -> 无法识别的音频格式")
            except Exception as e:
                self.log_signal.emit(f"[失败] {filename} -> 错误原因: {str(e)}")

            # 更新进度条
            self.progress_signal.emit(i + 1, total_files)

        self.log_signal.emit("\n🎉 所有文件处理完毕！")
        self.finished_signal.emit()


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🎵 音乐标签极净清除工具")
        self.resize(600, 500)
        self.setAcceptDrops(True) # 开启全局拖拽支持
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # 拖拽提示区域
        self.drop_label = QLabel("将一个或多个音乐文件/文件夹\n拖拽到此处\n(支持 MP3, WAV, FLAC)")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setFont(QFont("Microsoft YaHei", 12))
        self.drop_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                background-color: #f9f9f9;
                padding: 40px;
                color: #555;
            }
        """)
        layout.addWidget(self.drop_label)

        # 按钮区域
        btn_layout = QHBoxLayout()
        
        self.btn_select_files = QPushButton("📁 选择文件")
        self.btn_select_files.setMinimumHeight(40)
        self.btn_select_files.clicked.connect(self.select_files)
        
        self.btn_select_folder = QPushButton("📂 选择文件夹")
        self.btn_select_folder.setMinimumHeight(40)
        self.btn_select_folder.clicked.connect(self.select_folder)
        
        btn_layout.addWidget(self.btn_select_files)
        btn_layout.addWidget(self.btn_select_folder)
        layout.addLayout(btn_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        # 日志输出区域
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setFont(QFont("Consolas", 10))
        self.log_area.setStyleSheet("background-color: #2b2b2b; color: #a9b7c6;")
        layout.addWidget(self.log_area)

        self.log_msg("等待导入文件...\n")

    def log_msg(self, msg):
        self.log_area.append(msg)
        # 自动滚动到底部
        self.log_area.verticalScrollBar().setValue(self.log_area.verticalScrollBar().maximum())

    def update_progress(self, current, total):
        percentage = int((current / total) * 100)
        self.progress_bar.setValue(percentage)
        self.progress_bar.setFormat(f"处理进度: {current}/{total} ({percentage}%)")

    def process_paths(self, paths):
        """启动后台线程处理传入的路径列表"""
        if not paths:
            return
            
        self.btn_select_files.setEnabled(False)
        self.btn_select_folder.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_area.clear()
        
        self.worker = TagCleanerWorker(paths)
        self.worker.log_signal.connect(self.log_msg)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.on_processing_finished)
        self.worker.start()

    def on_processing_finished(self):
        self.btn_select_files.setEnabled(True)
        self.btn_select_folder.setEnabled(True)

    # --- 按钮事件 ---
    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择音频文件", "", "Audio Files (*.mp3 *.wav *.flac)"
        )
        if files:
            self.process_paths(files)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.process_paths([folder])

    # --- 拖拽事件支持 ---
    def dragEnterEvent(self, event):
        # 检查拖入的是否是文件或文件夹
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_label.setStyleSheet("""
                QLabel {
                    border: 2px solid #4CAF50;
                    border-radius: 10px;
                    background-color: #e8f5e9;
                    padding: 40px;
                    color: #2E7D32;
                }
            """)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        # 恢复默认样式
        self.drop_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                background-color: #f9f9f9;
                padding: 40px;
                color: #555;
            }
        """)

    def dropEvent(self, event):
        self.dragLeaveEvent(event) # 恢复样式
        urls = event.mimeData().urls()
        paths = [url.toLocalFile() for url in urls if url.isLocalFile()]
        if paths:
            self.process_paths(paths)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # 设置应用全局字体
    font = QFont("Microsoft YaHei", 10)
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
