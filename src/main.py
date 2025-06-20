import sys
import os

# 添加当前目录的父目录到Python路径
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from src.ui.main_window import MainWindow

def main():
    # 创建应用程序
    app = QApplication(sys.argv)
    
    # 设置全局字体
    font = QFont("Microsoft YaHei UI", 9)
    app.setFont(font)
    
    # 设置应用程序属性
    app.setApplicationName("PPT布局工具")
    app.setOrganizationName("PPT工具")
    
    # 使用Qt的新式风格
    app.setStyle("Fusion")
    
    # 创建并显示主窗口
    window = MainWindow()
    window.show()
    
    # 运行应用程序
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 