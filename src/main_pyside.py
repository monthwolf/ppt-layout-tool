"""
PySide6版本的程序入口点
如果PyQt6版本无法运行，可以使用此版本
"""
import sys
import os

# 添加当前目录的父目录到Python路径
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

try:
    from PySide6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget
    from PySide6.QtWidgets import QPushButton, QSpinBox, QLabel, QFileDialog, QScrollArea, QGroupBox
    from PySide6.QtWidgets import QDoubleSpinBox, QMessageBox, QSizePolicy, QFrame, QGridLayout
    from PySide6.QtCore import Qt, QRectF, QSize
    from PySide6.QtGui import QPixmap, QPainter, QPen, QColor, QFont
    
    USE_PYSIDE = True
    print("使用PySide6库")
except ImportError:
    print("PySide6未安装，尝试导入PyQt6...")
    from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget
    from PyQt6.QtWidgets import QPushButton, QSpinBox, QLabel, QFileDialog, QScrollArea, QGroupBox
    from PyQt6.QtWidgets import QDoubleSpinBox, QMessageBox, QSizePolicy, QFrame, QGridLayout
    from PyQt6.QtCore import Qt, QRectF, QSize
    from PyQt6.QtGui import QPixmap, QPainter, QPen, QColor, QFont
    
    USE_PYSIDE = False
    print("使用PyQt6库")

# 导入本地模块，注意UI需要根据UI库分别适配
from src.utils.ppt_processor import PPTProcessor
from src.utils.layout_calculator import LayoutCalculator

# 定义简单样式
COLORS = {
    'primary': '#2979FF',
    'background': '#FAFAFA',
    'surface': '#FFFFFF',
    'text_primary': '#212121',
}

STYLESHEET = f"""
QWidget {{
    font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif;
    color: {COLORS['text_primary']};
    background-color: {COLORS['background']};
}}

QPushButton {{
    background-color: {COLORS['primary']};
    color: white;
    border: none;
    border-radius: 4px;
    padding: 8px 16px;
    min-width: 80px;
}}
"""

class SimpleMainWindow(QMainWindow):
    """简化版主窗口，适用于PyQt6和PySide6"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PPT布局工具 (兼容版)")
        self.setMinimumSize(800, 600)
        
        self.ppt_processor = PPTProcessor()
        self.layout_calculator = LayoutCalculator()
        
        self.current_ppt_path = None
        self.slide_images = []
        self.layout_config = {
            "columns": 2,
            "page_width": 210,
            "page_height": 297,
            "margin_left": 10,
            "margin_top": 10,
            "margin_right": 10,
            "margin_bottom": 10,
            "h_spacing": 5,
            "v_spacing": 5,
        }
        
        self.init_ui()
    
    def init_ui(self):
        """初始化UI界面"""
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        
        # 文件选择区域
        file_layout = QHBoxLayout()
        self.select_ppt_btn = QPushButton("选择PPT文件")
        self.select_ppt_btn.clicked.connect(self.select_ppt_file)
        file_layout.addWidget(self.select_ppt_btn)
        
        self.file_label = QLabel("未选择文件")
        file_layout.addWidget(self.file_label)
        file_layout.addStretch(1)
        
        main_layout.addLayout(file_layout)
        
        # 布局设置区域
        settings_group = QGroupBox("布局设置")
        settings_layout = QGridLayout(settings_group)
        
        settings_layout.addWidget(QLabel("每行PPT数量:"), 0, 0)
        self.columns_spin = QSpinBox()
        self.columns_spin.setRange(1, 10)
        self.columns_spin.setValue(self.layout_config["columns"])
        self.columns_spin.valueChanged.connect(self.update_layout)
        settings_layout.addWidget(self.columns_spin, 0, 1)
        
        settings_layout.addWidget(QLabel("水平间距(mm):"), 1, 0)
        self.h_spacing_spin = QDoubleSpinBox()
        self.h_spacing_spin.setRange(0, 50)
        self.h_spacing_spin.setValue(self.layout_config["h_spacing"])
        self.h_spacing_spin.valueChanged.connect(self.update_spacing)
        settings_layout.addWidget(self.h_spacing_spin, 1, 1)
        
        settings_layout.addWidget(QLabel("垂直间距(mm):"), 2, 0)
        self.v_spacing_spin = QDoubleSpinBox()
        self.v_spacing_spin.setRange(0, 50)
        self.v_spacing_spin.setValue(self.layout_config["v_spacing"])
        self.v_spacing_spin.valueChanged.connect(self.update_spacing)
        settings_layout.addWidget(self.v_spacing_spin, 2, 1)
        
        main_layout.addWidget(settings_group)
        
        # 预览区域
        self.preview_info = QLabel("请先选择PPT文件并设置布局")
        main_layout.addWidget(self.preview_info)
        
        # 按钮区域
        btn_layout = QHBoxLayout()
        btn_layout.addStretch(1)
        
        self.preview_btn = QPushButton("计算并预览布局")
        self.preview_btn.clicked.connect(self.preview_layout)
        self.preview_btn.setEnabled(False)
        btn_layout.addWidget(self.preview_btn)
        
        self.process_btn = QPushButton("导出PDF")
        self.process_btn.clicked.connect(self.process_ppt)
        self.process_btn.setEnabled(False)
        btn_layout.addWidget(self.process_btn)
        
        main_layout.addLayout(btn_layout)
        
        # 预览滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        self.preview_widget = QWidget()
        self.preview_layout = QVBoxLayout(self.preview_widget)
        self.preview_layout.setAlignment(Qt.AlignmentFlag.AlignCenter if not USE_PYSIDE else Qt.AlignCenter)
        scroll_area.setWidget(self.preview_widget)
        
        main_layout.addWidget(scroll_area)
        
        self.setCentralWidget(central_widget)
    
    def select_ppt_file(self):
        """选择PPT文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择PPT文件", "", "PowerPoint文件 (*.pptx *.ppt)"
        )
        
        if file_path:
            self.current_ppt_path = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.preview_btn.setEnabled(True)
            self.process_btn.setEnabled(False)
            self.slide_images = self.ppt_processor.convert_ppt_to_images(file_path)
            self.preview_info.setText(f"成功导入 {len(self.slide_images)} 张PPT幻灯片")
    
    def update_layout(self):
        """更新布局设置"""
        self.layout_config["columns"] = self.columns_spin.value()
    
    def update_spacing(self):
        """更新间距设置"""
        self.layout_config["h_spacing"] = self.h_spacing_spin.value()
        self.layout_config["v_spacing"] = self.v_spacing_spin.value()
    
    def preview_layout(self):
        """预览布局"""
        if not self.slide_images:
            return
        
        # 清除当前预览
        for i in reversed(range(self.preview_layout.count())):
            widget = self.preview_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        
        # 计算布局
        layout_result = self.layout_calculator.calculate_layout(
            self.slide_images, self.layout_config
        )
        
        # 显示结果
        result_text = f"布局结果: 每页 {layout_result['rows']} 行 × {layout_result['columns']} 列, "
        result_text += f"每个PPT尺寸: {layout_result['item_width']:.1f} × {layout_result['item_height']:.1f} mm"
        self.preview_info.setText(result_text)
        
        # 启用导出按钮
        self.process_btn.setEnabled(True)
    
    def process_ppt(self):
        """处理PPT并导出PDF"""
        if not self.slide_images:
            return
        
        # 选择输出文件
        output_path, _ = QFileDialog.getSaveFileName(
            self, "保存PDF文件", "", "PDF文件 (*.pdf)"
        )
        
        if not output_path:
            return
            
        # 使用当前布局设置处理PPT并生成PDF
        layout_result = self.layout_calculator.calculate_layout(
            self.slide_images, self.layout_config
        )
        
        success = self.ppt_processor.generate_pdf(
            self.slide_images, 
            output_path,
            layout_result, 
            self.layout_config
        )
        
        if success:
            QMessageBox.information(self, "成功", f"PDF已成功保存到:\n{output_path}")
        else:
            QMessageBox.critical(self, "错误", "生成PDF时出错")

def main():
    """主函数"""
    app = QApplication(sys.argv)
    
    # 设置样式表
    app.setStyleSheet(STYLESHEET)
    
    window = SimpleMainWindow()
    window.show()
    
    sys.exit(app.exec() if not USE_PYSIDE else app.exec_())

if __name__ == "__main__":
    main() 