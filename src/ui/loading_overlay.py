import math
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QFrame, QLabel
from PyQt6.QtCore import Qt, QSize, QPropertyAnimation, QEasingCurve, QByteArray
from PyQt6.QtGui import QPainter, QColor, QPen, QBrush

from .spinner_widget import SpinnerWidget

class LoadingOverlay(QWidget):
    """一个经过美化的、带自定义加载动画的覆盖层。"""
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

        self.opacity_animation = QPropertyAnimation(self, QByteArray(b"windowOpacity"))
        self.opacity_animation.setDuration(300)
        self.opacity_animation.setEasingCurve(QEasingCurve.Type.InOutQuad)

    def set_text(self, text):
        self.loading_text.setText(text)

    def show(self):
        if self.parentWidget():
            self.resize(self.parentWidget().size())
        
        self.setWindowOpacity(0.0)
        super().show()
        
        self.spinner.start()
        self.opacity_animation.setStartValue(0.0)
        self.opacity_animation.setEndValue(1.0)
        self.opacity_animation.start()

    def hide(self):
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
