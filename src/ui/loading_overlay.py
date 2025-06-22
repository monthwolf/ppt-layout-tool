import math
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QFrame, QLabel, QProgressBar
from PyQt6.QtCore import Qt, QSize, QPropertyAnimation, QEasingCurve, QByteArray
from PyQt6.QtGui import QPainter, QColor, QPen, QBrush

from .spinner_widget import SpinnerWidget

class LoadingOverlay(QWidget):
    """一个经过美化的、带自定义加载动画和进度条的覆盖层。"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setVisible(False)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)

        self.container = QFrame(self)
        self.container.setObjectName("loadingContainer")
        self.container.setStyleSheet("""
            #loadingContainer {
                background-color: rgba(252, 252, 252, 0.95);
                border: 1px solid #EAEAEA;
                border-radius: 18px;
                padding: 35px;
            }
        """)
        container_layout = QVBoxLayout(self.container)
        container_layout.setSpacing(20)

        self.spinner = SpinnerWidget()
        self.spinner.setFixedSize(50, 50)
        container_layout.addWidget(self.spinner, 0, Qt.AlignmentFlag.AlignCenter)

        self.loading_text = QLabel("正在处理...")
        self.loading_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_text.setStyleSheet("""
            font-size: 15px; 
            font-weight: 500; 
            color: #555555;
            background: transparent;
        """)
        container_layout.addWidget(self.loading_text)
        
        # 添加进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedWidth(250)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #E0E0E0;
                border-radius: 5px;
                text-align: center;
                height: 20px;
                background-color: #F8F8F8;
            }
            
            QProgressBar::chunk {
                background-color: #4A90E2;
                border-radius: 5px;
            }
        """)
        self.progress_bar.setVisible(False)  # 默认隐藏进度条
        container_layout.addWidget(self.progress_bar, 0, Qt.AlignmentFlag.AlignCenter)
        
        # 添加进度文本
        self.progress_text = QLabel("")
        self.progress_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_text.setStyleSheet("""
            font-size: 13px;
            color: #777777;
            background: transparent;
        """)
        self.progress_text.setVisible(False)  # 默认隐藏进度文本
        container_layout.addWidget(self.progress_text)

        layout.addWidget(self.container, 0, Qt.AlignmentFlag.AlignCenter)
        # 创建属性动画
        self.opacity_animation = QPropertyAnimation(self, b"windowOpacity")
        self.opacity_animation.setDuration(300)
        self.opacity_animation.setEasingCurve(QEasingCurve.Type.InOutQuad)

    def set_text(self, text):
        """设置加载文本"""
        self.loading_text.setText(text)
    
    def set_progress(self, value, max_value=100, text=None):
        """
        设置进度条值和文本
        
        Args:
            value: 当前进度值
            max_value: 最大进度值
            text: 进度文本，如果为None则自动生成
        """
        # 显示进度条和进度文本
        self.progress_bar.setVisible(True)
        self.progress_text.setVisible(True)
        
        # 设置进度条范围和当前值
        self.progress_bar.setRange(0, max_value)
        self.progress_bar.setValue(value)
        
        # 设置进度文本
        if text is None:
            percentage = int(value / max_value * 100) if max_value > 0 else 0
            text = f"当前进度: {percentage}% ({value}/{max_value})"
        self.progress_text.setText(text)
    
    def hide_progress(self):
        """隐藏进度条和进度文本"""
        self.progress_bar.setVisible(False)
        self.progress_text.setVisible(False)

    def show(self):
        """显示加载覆盖层"""
        if self.parentWidget():
            self.resize(self.parentWidget().size())
        
        self.setWindowOpacity(0.0)
        super().show()
        
        self.spinner.start()
        self.opacity_animation.setStartValue(0.0)
        self.opacity_animation.setEndValue(1.0)
        self.opacity_animation.start()

    def hide(self):
        """隐藏加载覆盖层"""
        # 重置进度条和进度文本
        self.progress_bar.setValue(0)
        self.progress_text.setText("")
        self.hide_progress()
        
        self.opacity_animation.setStartValue(1.0)
        self.opacity_animation.setEndValue(0.0)
        # 使用lambda表达式确保只在需要时连接，避免重复连接
        self.opacity_animation.finished.connect(self._on_hide_finished)
        self.opacity_animation.start()

    def _on_hide_finished(self):
        super().hide()
        self.spinner.stop()
        try:
            self.opacity_animation.finished.disconnect(self._on_hide_finished)
        except TypeError:
            pass # 如果已经断开连接，则忽略错误
