
import sys
import os
import subprocess
import time
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                            QHBoxLayout, QPushButton, QLabel, QFileDialog,
                            QProgressBar, QMessageBox)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QIcon, QColor, QPalette, QFont

class ConvertThread(QThread):
    progress = pyqtSignal(int)
    current_file = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, input_folder, output_folder):
        super().__init__()
        self.input_folder = input_folder
        self.output_folder = output_folder
        self.is_paused = False
        self.is_stopped = False
        self.word_app = None

    def pause(self):
        self.is_paused = True

    def resume(self):
        self.is_paused = False

    def stop(self):
        self.is_stopped = True
        if self.word_app:
            try:
                subprocess.run(['osascript', '-e', 'tell application "Microsoft Word" to quit'], check=True)
            except:
                pass

    def run(self):
        try:
            files = [f for f in os.listdir(self.input_folder)
                    if f.endswith(('.docx', '.doc'))]
            total_files = len(files)
            
            if total_files == 0:
                self.error.emit('未找到可转换的Word文档')
                return
            
            self.current_file.emit(f'已发现 {total_files} 个Word文件')
            converted_count = 0
            
            # 启动Word并保持运行
            init_script = '''
            tell application "Microsoft Word"
                launch
                set bounds of window 1 to {100, 100, 600, 800}
                set visible to true
            end tell
            '''
            subprocess.run(['osascript', '-e', init_script], check=True)
            self.word_app = True
            
            for i, filename in enumerate(files):
                if self.is_stopped:
                    self.error.emit('转换已停止')
                    return
                    
                while self.is_paused:
                    time.sleep(0.1)
                    if self.is_stopped:
                        self.error.emit('转换已停止')
                        return
                        
                try:
                    input_path = os.path.join(self.input_folder, filename)
                    output_path = os.path.join(self.output_folder, 
                                             os.path.splitext(filename)[0] + '.pdf')
                    
                    applescript = f'''
                    tell application "Microsoft Word"
                        set bounds of window 1 to {100, 100, 600, 800}
                        set visible to true
                        open "{input_path}"
                        set activeDoc to active document
                        save as activeDoc file name "{output_path}" file format format PDF
                        close activeDoc saving no
                    end tell
                    '''
                    
                    converted_count += 1
                    self.current_file.emit(f'已发现 {total_files} 个Word文件，已成功转换 {converted_count} 个')
                    subprocess.run(['osascript', '-e', applescript], check=True)
                    self.progress.emit(int((i + 1) / total_files * 100))
                except Exception as e:
                    self.error.emit(f'处理文件 {filename} 时出错：{str(e)}')
                    return
            
            # 完成后关闭Word
            if self.word_app:
                subprocess.run(['osascript', '-e', 'tell application "Microsoft Word" to quit'], check=True)
            
            self.finished.emit()
        except Exception as e:
            self.error.emit(f'转换过程出错：{str(e)}')



class Word2PDFApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.check_office_installation()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Word to PDF Converter')
        self.setGeometry(100, 100, 500, 350)
        self.setMinimumSize(500, 350)

        main_widget = QWidget()
        layout = QVBoxLayout()

        # Input folder selection
        input_layout = QHBoxLayout()
        self.input_label = QLabel('选择输入文件夹：未选择')
        self.input_label.setStyleSheet("color: #333333; font-size: 14px;")
        input_button = QPushButton('选择输入文件夹')
        input_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        input_button.clicked.connect(self.select_input_folder)
        input_layout.addWidget(self.input_label)
        input_layout.addWidget(input_button)

        # Output folder selection
        output_layout = QHBoxLayout()
        self.output_label = QLabel('选择输出文件夹：未选择')
        self.output_label.setStyleSheet("color: #333333; font-size: 14px;")
        output_button = QPushButton('选择输出文件夹')
        output_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-size: 14px;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #1e88e5;
            }
        """)
        output_button.clicked.connect(self.select_output_folder)
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(output_button)

        # Current file label
        self.current_file_label = QLabel('准备就绪')
        self.current_file_label.setStyleSheet("color: #666666; font-size: 12px;")
        self.current_file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 4px;
                text-align: center;
                background-color: #E0E0E0;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 4px;
            }
        """)

        # Control buttons
        control_layout = QHBoxLayout()
        
        # Convert button
        self.convert_button = QPushButton('开始转换')
        self.convert_button.setStyleSheet("""
            QPushButton {
                background-color: #F44336;
                color: white;
                font-size: 14px;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #e53935;
            }
        """)
        self.convert_button.clicked.connect(self.start_conversion)
        
        # Pause button
        self.pause_button = QPushButton('暂停')
        self.pause_button.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                font-size: 14px;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #fb8c00;
            }
        """)
        self.pause_button.clicked.connect(self.toggle_pause)
        self.pause_button.setEnabled(False)
        
        # Stop button
        self.stop_button = QPushButton('停止')
        self.stop_button.setStyleSheet("""
            QPushButton {
                background-color: #9E9E9E;
                color: white;
                font-size: 14px;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #757575;
            }
        """)
        self.stop_button.clicked.connect(self.stop_conversion)
        self.stop_button.setEnabled(False)

        control_layout.addWidget(self.convert_button)
        control_layout.addWidget(self.pause_button)
        control_layout.addWidget(self.stop_button)

        # Add widgets to layout
        layout.addLayout(input_layout)
        layout.addLayout(output_layout)
        layout.addWidget(self.current_file_label)
        layout.addWidget(self.progress_bar)
        layout.addLayout(control_layout)
        
        # Add some padding
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        main_widget.setLayout(layout)
        self.setCentralWidget(main_widget)

    def select_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, '选择输入文件夹')
        if folder:
            self.input_folder = folder
            self.input_label.setText(f'选择输入文件夹：{folder}')

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, '选择输出文件夹')
        if folder:
            self.output_folder = folder
            self.output_label.setText(f'选择输出文件夹：{folder}')

    def check_office_installation(self):
        try:
            # 尝试运行权限授予脚本，但不显示对话框
            grant_script = os.path.join(os.path.dirname(__file__), 'grant_access.scpt')
            if os.path.exists(grant_script):
                try:
                    # 静默运行权限检查
                    subprocess.run(['osascript', '-e', 'tell application "System Events" to get name of application processes'], check=True)
                except subprocess.CalledProcessError:
                    print('警告：无法获取系统事件权限')
        except Exception as e:
            QMessageBox.critical(
                self,
                '错误',
                f'初始化程序时发生未知错误：{str(e)}\n\n程序将退出。'
            )
            sys.exit(1)

    def start_conversion(self):
        if not hasattr(self, 'input_folder') or not hasattr(self, 'output_folder'):
            QMessageBox.warning(self, '提示', '请先选择输入和输出文件夹。')
            return

        self.convert_thread = ConvertThread(self.input_folder, self.output_folder)
        self.convert_thread.progress.connect(self.update_progress)
        self.convert_thread.current_file.connect(self.update_current_file)
        self.convert_thread.finished.connect(self.conversion_finished)
        self.convert_thread.error.connect(self.show_error)
        self.convert_thread.start()
        
        # 更新按钮状态
        self.convert_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.stop_button.setEnabled(True)

    def toggle_pause(self):
        if not hasattr(self, 'convert_thread'):
            return
            
        if self.convert_thread.is_paused:
            self.convert_thread.resume()
            self.pause_button.setText('暂停')
            self.pause_button.setStyleSheet("""
                QPushButton {
                    background-color: #FF9800;
                    color: white;
                    font-size: 14px;
                    border: none;
                    border-radius: 4px;
                    padding: 8px 16px;
                }
                QPushButton:hover {
                    background-color: #fb8c00;
                }
            """)
        else:
            self.convert_thread.pause()
            self.pause_button.setText('继续')
            self.pause_button.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    font-size: 14px;
                    border: none;
                    border-radius: 4px;
                    padding: 8px 16px;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
            """)

    def stop_conversion(self):
        if not hasattr(self, 'convert_thread'):
            return
            
        self.convert_thread.stop()
        self.pause_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.convert_button.setEnabled(True)
        self.current_file_label.setText('转换已停止')

    def conversion_finished(self):
        self.progress_bar.setValue(100)
        self.current_file_label.setText('转换完成！')
        self.convert_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        QMessageBox.information(self, '完成', '所有文件已转换完成！')

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_current_file(self, filename):
        self.current_file_label.setText(filename)

    def show_error(self, error_message):
        QMessageBox.critical(self, '错误', error_message)

if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        app.setStyle('Fusion')
        
        # 设置系统默认字体
        app.setFont(app.font())
        
        # 设置应用图标
        icon_path = os.path.join(os.path.dirname(__file__), 'icon.svg')
        if os.path.exists(icon_path):
            app.setWindowIcon(QIcon(icon_path))
        
        # 设置应用主题
        palette = app.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(255, 255, 255))
        app.setPalette(palette)
        
        # 创建并显示主窗口
        window = Word2PDFApp()
        window.setWindowIcon(QIcon(icon_path))
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        QMessageBox.critical(None, '错误', f'程序启动时发生错误：{str(e)}')
        sys.exit(1)

