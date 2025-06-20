from PyQt6.QtWidgets import QWidget
from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QPainter, QColor, QPen

class SpinnerWidget(QWidget):
    """一个自定义绘制的、平滑的加载旋转动画。"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._angle = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._update_angle)
        self.setFixedSize(28, 28) # 调整为适合步骤指示器的大小

    def _update_angle(self):
        self._angle = (self._angle + 10) % 360
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        pen = QPen(QColor("#4A90E2"), 3) # 调整画笔粗细
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(pen)
        
        painter.drawArc(
            self.rect().adjusted(2, 2, -2, -2), # 调整边距
            self._angle * 16,
            90 * 16
        )
    
    def start(self):
        self._timer.start(15)

    def stop(self):
        self._timer.stop()
