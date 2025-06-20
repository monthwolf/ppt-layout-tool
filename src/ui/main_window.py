from PyQt6.QtWidgets import (QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, 
                           QPushButton, QSpinBox, QLabel, QFileDialog, 
                           QScrollArea, QGroupBox, QDoubleSpinBox, QMessageBox,
                           QSizePolicy, QFrame, QGridLayout, QProgressBar,
                           QStatusBar, QToolButton, QTabWidget, QWizard, 
                           QWizardPage, QStackedWidget, QRadioButton, QButtonGroup,
                           QCheckBox, QTextEdit, QApplication)
from PyQt6.QtCore import Qt, QRectF, QSize, QThread, pyqtSignal
from PyQt6.QtGui import QPixmap, QPainter, QPen, QColor, QFont, QIcon, QMovie, QPainterPath

import os
from pptx import Presentation
from PIL import Image
import io

from src.utils.ppt_processor import PPTProcessor
from src.utils.layout_calculator import LayoutCalculator
from src.ui.styles import STYLESHEET, COLORS, WELCOME_TEXT, STEPS_GUIDE

class LoadingOverlay(QWidget):
    """一个现代化的加载覆盖层，提供视觉反馈"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setVisible(False)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # 加载动画容器
        self.container = QFrame()
        self.container.setObjectName("loadingContainer")
        self.container.setStyleSheet(f"""
            #loadingContainer {{
                background-color: rgba(0, 0, 0, 0.8);
                border-radius: 15px;
                padding: 30px;
                color: white;
            }}
        """)
        container_layout = QVBoxLayout(self.container)
        container_layout.setSpacing(20)

        # GIF动画
        self.loading_animation = QLabel()
        loading_gif_path = os.path.join("resources", "loading.gif")
        if os.path.exists(loading_gif_path):
            self.movie = QMovie(loading_gif_path)
            self.movie.setScaledSize(QSize(64, 64))
            self.loading_animation.setMovie(self.movie)
        else:
            self.loading_animation.setText("B") # Fallback
        container_layout.addWidget(self.loading_animation, 0, Qt.AlignmentFlag.AlignCenter)

        # 加载文本
        self.loading_text = QLabel("正在处理...")
        self.loading_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_text.setStyleSheet("font-size: 16px; font-weight: bold;")
        container_layout.addWidget(self.loading_text)
        
        layout.addWidget(self.container)

    def showEvent(self, event):
        super().showEvent(event)
        if hasattr(self, 'movie'):
            self.movie.start()

    def hideEvent(self, event):
        super().hideEvent(event)
        if hasattr(self, 'movie'):
            self.movie.stop()

    def set_text(self, text):
        self.loading_text.setText(text)

    def show(self):
        if self.parentWidget():
            self.resize(self.parentWidget().size())
        super().show()

class Worker(QThread):
    """通用工作线程，用于执行耗时操作"""
    finished = pyqtSignal(object)
    error = pyqtSignal(Exception)

    def __init__(self, function, *args, **kwargs):
        super().__init__()
        self.function = function
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            result = self.function(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(e)

class StepIndicator(QFrame):
    """步骤指示器组件"""
    def __init__(self, steps, parent=None):
        super().__init__(parent)
        self.steps = steps
        self.current_step = 0
        
        layout = QHBoxLayout(self)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 10, 0, 10)
        
        self.step_widgets = []
        for i, step_text in enumerate(steps):
            # Step Container
            step_container = QWidget()
            step_layout = QVBoxLayout(step_container)
            step_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            step_layout.setContentsMargins(5, 0, 5, 0)

            # Icon/Number Label
            icon_label = QLabel()
            icon_label.setFixedSize(32, 32)
            icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            
            # Text Label
            text_label = QLabel(step_text)
            text_label.setObjectName(f"stepLabel{i}")
            
            step_layout.addWidget(icon_label, 0, Qt.AlignmentFlag.AlignCenter)
            step_layout.addWidget(text_label, 0, Qt.AlignmentFlag.AlignCenter)
            
            layout.addWidget(step_container)
            self.step_widgets.append({'container': step_container, 'icon': icon_label, 'text': text_label})
            
            if i < len(steps) - 1:
                line = QFrame()
                line.setFrameShape(QFrame.Shape.HLine)
                line.setFixedHeight(2)
                line.setStyleSheet(f"background-color: {COLORS['divider']};")
                line.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
                layout.addWidget(line)
                
        # 准备加载动画
        # 注意: 需要一个loading.gif文件放在resources目录下
        loading_gif_path = os.path.join("resources", "loading.gif")
        self.loading_movie = QMovie(loading_gif_path)
        self.loading_movie.setScaledSize(QSize(28, 28))

    def set_current_step(self, step):
        if 0 <= step < len(self.steps):
            self.current_step = step
            
            for i, widget_group in enumerate(self.step_widgets):
                icon_label = widget_group['icon']
                text_label = widget_group['text']
                
                # 停止之前的动画
                if icon_label.movie():
                    icon_label.movie().stop()
                    icon_label.setMovie(None)

                if i < step:
                    # 已完成步骤: 显示勾选图标
                    pixmap = QPixmap(32, 32)
                    pixmap.fill(Qt.GlobalColor.transparent)
                    p = QPainter(pixmap)
                    p.setRenderHint(QPainter.RenderHint.Antialiasing)
                    p.setBrush(QColor(COLORS['success']))
                    p.setPen(Qt.GlobalColor.transparent)
                    p.drawEllipse(0, 0, 32, 32)
                    
                    pen = QPen(QColor("white"), 2)
                    p.setPen(pen)
                    path = QPainterPath()
                    path.moveTo(9, 16)
                    path.lineTo(14, 21)
                    path.lineTo(23, 12)
                    p.drawPath(path)
                    p.end()
                    
                    icon_label.setPixmap(pixmap)
                    text_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-weight: normal;")
                    
                elif i == step:
                    # 当前步骤: 显示加载动画
                    if self.loading_movie.isValid():
                        icon_label.setMovie(self.loading_movie)
                        self.loading_movie.start()
                    else:
                        # Fallback to number if GIF not found
                        icon_label.setText(str(i + 1))
                        icon_label.setStyleSheet(f"""
                            background-color: {COLORS['primary']};
                            border-radius: 16px;
                            color: white;
                            font-weight: bold;
                        """)
                    text_label.setStyleSheet(f"color: {COLORS['primary']}; font-weight: bold;")
                    
                else:
                    # 未完成步骤: 显示数字
                    icon_label.setText(str(i + 1))
                    icon_label.setStyleSheet(f"""
                        background-color: {COLORS['divider']};
                        border-radius: 16px;
                        color: {COLORS['text_secondary']};
                        font-weight: normal;
                    """)
                    text_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-weight: normal;")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PPT布局工具")
        self.setMinimumSize(1100, 800)
        
        self.ppt_processor = PPTProcessor()
        self.layout_calculator = LayoutCalculator()
        
        self.current_ppt_path = None
        self.slide_images = []
        self.layout_config = {
            "columns": 2,  # 默认每行2列
            "page_width": 210,  # A4宽度(mm)
            "page_height": 297,  # A4高度(mm)
            "margin_left": 10,
            "margin_top": 10,
            "margin_right": 10,
            "margin_bottom": 10,
            "h_spacing": 5,  # 水平间距(mm)
            "v_spacing": 5,  # 垂直间距(mm)
            "is_landscape": True,  # 默认为横向A4
            "show_ppt_numbers": True,  # 显示PPT定位页码
            "show_page_numbers": True,  # 显示纸张页码
        }
        
        self.current_step = 0
        
        # 应用样式表
        self.setStyleSheet(STYLESHEET)
        
        self.init_ui()
        self.init_loading_overlay()
        
        # 显示欢迎界面
        self.show_welcome_screen()
        
        # 创建状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("准备就绪")
    
    def init_loading_overlay(self):
        """初始化加载覆盖层"""
        self.loading_overlay = LoadingOverlay(self)
    
    def resizeEvent(self, event):
        """确保覆盖层始终与主窗口大小一致"""
        super().resizeEvent(event)
        if hasattr(self, 'loading_overlay'):
            self.loading_overlay.resize(event.size())
    
    def init_ui(self):
        """初始化UI界面"""
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        
        # 添加步骤指示器
        self.step_indicator = StepIndicator(STEPS_GUIDE)
        main_layout.addWidget(self.step_indicator)
        
        # 创建堆叠式部件用于不同步骤的界面
        self.stacked_widget = QStackedWidget()
        
        # 创建各步骤页面
        self.create_step1_page()  # 文件选择页面
        self.create_step2_page()  # 布局设置页面
        self.create_step3_page()  # 预览页面
        self.create_step4_page()  # 导出页面
        self.create_step5_page()  # AI索引页面
        
        main_layout.addWidget(self.stacked_widget)
        
        # 创建导航按钮
        nav_layout = QHBoxLayout()
        
        # 占位空间
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        nav_layout.addWidget(spacer)
        
        # 上一步按钮
        self.prev_btn = QPushButton("< 上一步")
        self.prev_btn.clicked.connect(self.go_to_prev_step)
        self.prev_btn.setEnabled(False)
        nav_layout.addWidget(self.prev_btn)
        
        # 下一步按钮
        self.next_btn = QPushButton("下一步 >")
        self.next_btn.setObjectName("accentButton")
        self.next_btn.clicked.connect(self.go_to_next_step)
        self.next_btn.setEnabled(False)  # 初始状态禁用，等待选择文件
        nav_layout.addWidget(self.next_btn)
        
        main_layout.addLayout(nav_layout)
        
        self.setCentralWidget(central_widget)
    
    def create_step1_page(self):
        """创建步骤1：选择PPT文件页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 20, 0, 0)
        layout.setSpacing(20)
        
        # 创建文件选择区域
        file_group = QGroupBox("选择PPT文件")
        file_layout = QVBoxLayout(file_group)
        
        # 选择PPT按钮
        file_btn_layout = QHBoxLayout()
        file_btn_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.select_ppt_btn = QPushButton("选择PPT文件")
        self.select_ppt_btn.setMinimumHeight(40)
        self.select_ppt_btn.clicked.connect(self.select_ppt_file)
        file_btn_layout.addWidget(self.select_ppt_btn)
        
        file_layout.addLayout(file_btn_layout)
        
        # 显示当前文件信息
        self.file_info = QLabel()
        self.file_info.setObjectName("infoLabel")
        self.file_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_info.setText("未选择任何文件")
        file_layout.addWidget(self.file_info)
        
        layout.addWidget(file_group)
        
        # 文件预览区域
        preview_group = QGroupBox("文件信息")
        preview_layout = QVBoxLayout(preview_group)
        
        self.file_preview = QLabel()
        self.file_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_preview.setText("选择PPT文件后将显示幻灯片数量及预览信息")
        preview_layout.addWidget(self.file_preview)
        
        layout.addWidget(preview_group)
        
        # 添加到堆叠式部件
        self.stacked_widget.addWidget(page)
    
    def create_step2_page(self):
        """创建步骤2：布局设置页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 20, 0, 0)
        layout.setSpacing(20)
        
        # 页面方向选择
        orientation_group = QGroupBox("页面方向")
        orientation_layout = QHBoxLayout(orientation_group)
        
        self.orientation_group = QButtonGroup(self)
        
        self.landscape_radio = QRadioButton("横向A4 (297×210mm)")
        self.portrait_radio = QRadioButton("纵向A4 (210×297mm)")
        
        # 默认选择横向
        self.landscape_radio.setChecked(self.layout_config["is_landscape"])
        self.portrait_radio.setChecked(not self.layout_config["is_landscape"])
        
        self.orientation_group.addButton(self.landscape_radio)
        self.orientation_group.addButton(self.portrait_radio)
        
        self.landscape_radio.clicked.connect(self.update_orientation)
        self.portrait_radio.clicked.connect(self.update_orientation)
        
        # 添加图标或预览图
        landscape_layout = QVBoxLayout()
        landscape_layout.addWidget(self.landscape_radio)
        
        portrait_layout = QVBoxLayout()
        portrait_layout.addWidget(self.portrait_radio)
        
        orientation_layout.addLayout(landscape_layout)
        orientation_layout.addSpacing(20)
        orientation_layout.addLayout(portrait_layout)
        
        layout.addWidget(orientation_group)
        
        # 布局设置区域
        settings_group = QGroupBox("布局设置")
        settings_layout = QGridLayout(settings_group)
        settings_layout.setColumnStretch(1, 1)
        settings_layout.setColumnStretch(3, 1)
        
        # 每行数量设置
        settings_layout.addWidget(QLabel("每行PPT数量:"), 0, 0)
        
        self.columns_spin = QSpinBox()
        self.columns_spin.setRange(1, 10)
        self.columns_spin.setValue(self.layout_config["columns"])
        self.columns_spin.valueChanged.connect(self.update_layout)
        settings_layout.addWidget(self.columns_spin, 0, 1)
        
        # 水平间距设置
        settings_layout.addWidget(QLabel("水平间距 (mm):"), 0, 2)
        
        self.h_spacing_spin = QDoubleSpinBox()
        self.h_spacing_spin.setRange(0, 50)
        self.h_spacing_spin.setValue(self.layout_config["h_spacing"])
        self.h_spacing_spin.setSingleStep(0.5)
        self.h_spacing_spin.valueChanged.connect(self.update_spacing)
        settings_layout.addWidget(self.h_spacing_spin, 0, 3)
        
        # 垂直间距设置
        settings_layout.addWidget(QLabel("垂直间距 (mm):"), 1, 0)
        
        self.v_spacing_spin = QDoubleSpinBox()
        self.v_spacing_spin.setRange(0, 50)
        self.v_spacing_spin.setValue(self.layout_config["v_spacing"])
        self.v_spacing_spin.setSingleStep(0.5)
        self.v_spacing_spin.valueChanged.connect(self.update_spacing)
        settings_layout.addWidget(self.v_spacing_spin, 1, 1)
        
        # 页边距设置
        settings_layout.addWidget(QLabel("页面左边距 (mm):"), 2, 0)
        
        self.margin_left_spin = QDoubleSpinBox()
        self.margin_left_spin.setRange(0, 50)
        self.margin_left_spin.setValue(self.layout_config["margin_left"])
        self.margin_left_spin.setSingleStep(0.5)
        self.margin_left_spin.valueChanged.connect(self.update_margins)
        settings_layout.addWidget(self.margin_left_spin, 2, 1)
        
        settings_layout.addWidget(QLabel("页面右边距 (mm):"), 2, 2)
        
        self.margin_right_spin = QDoubleSpinBox()
        self.margin_right_spin.setRange(0, 50)
        self.margin_right_spin.setValue(self.layout_config["margin_right"])
        self.margin_right_spin.setSingleStep(0.5)
        self.margin_right_spin.valueChanged.connect(self.update_margins)
        settings_layout.addWidget(self.margin_right_spin, 2, 3)
        
        settings_layout.addWidget(QLabel("页面上边距 (mm):"), 3, 0)
        
        self.margin_top_spin = QDoubleSpinBox()
        self.margin_top_spin.setRange(0, 50)
        self.margin_top_spin.setValue(self.layout_config["margin_top"])
        self.margin_top_spin.setSingleStep(0.5)
        self.margin_top_spin.valueChanged.connect(self.update_margins)
        settings_layout.addWidget(self.margin_top_spin, 3, 1)
        
        settings_layout.addWidget(QLabel("页面下边距 (mm):"), 3, 2)
        
        self.margin_bottom_spin = QDoubleSpinBox()
        self.margin_bottom_spin.setRange(0, 50)
        self.margin_bottom_spin.setValue(self.layout_config["margin_bottom"])
        self.margin_bottom_spin.setSingleStep(0.5)
        self.margin_bottom_spin.valueChanged.connect(self.update_margins)
        settings_layout.addWidget(self.margin_bottom_spin, 3, 3)
        
        # 添加显示页码选项
        settings_layout.addWidget(QLabel("页码设置:"), 4, 0)
        
        # 添加PPT定位页码选项
        self.show_ppt_numbers_check = QCheckBox("显示PPT定位页码")
        self.show_ppt_numbers_check.setChecked(self.layout_config["show_ppt_numbers"])
        self.show_ppt_numbers_check.stateChanged.connect(self.update_page_numbers)
        settings_layout.addWidget(self.show_ppt_numbers_check, 4, 1)
        
        # 添加纸张页码选项
        self.show_page_numbers_check = QCheckBox("显示纸张页码")
        self.show_page_numbers_check.setChecked(self.layout_config["show_page_numbers"])
        self.show_page_numbers_check.stateChanged.connect(self.update_page_numbers)
        settings_layout.addWidget(self.show_page_numbers_check, 4, 3)
        
        layout.addWidget(settings_group)
        
        # 提示信息
        hint_label = QLabel("设置好布局参数后，点击「下一步」查看预览效果")
        hint_label.setObjectName("infoLabel")
        hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(hint_label)
        
        # 添加到堆叠式部件
        self.stacked_widget.addWidget(page)
    
    def create_step3_page(self):
        """创建步骤3：预览页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 20, 0, 0)
        layout.setSpacing(20)
        
        # 预览信息区域
        info_group = QGroupBox("布局信息")
        info_layout = QVBoxLayout(info_group)
        
        self.preview_info = QLabel("计算中...")
        self.preview_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        info_layout.addWidget(self.preview_info)
        
        # 预览刷新按钮
        refresh_layout = QHBoxLayout()
        refresh_layout.setAlignment(Qt.AlignmentFlag.AlignRight)
        
        self.refresh_preview_btn = QPushButton("刷新预览")
        self.refresh_preview_btn.clicked.connect(self.refresh_preview)
        refresh_layout.addWidget(self.refresh_preview_btn)
        
        info_layout.addLayout(refresh_layout)
        
        layout.addWidget(info_group)
        
        # 预览区域
        preview_group = QGroupBox("布局预览")
        preview_layout = QVBoxLayout(preview_group)
        
        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        self.preview_widget = QWidget()
        self.preview_layout = QVBoxLayout(self.preview_widget)
        self.preview_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        scroll_area.setWidget(self.preview_widget)
        
        preview_layout.addWidget(scroll_area)
        
        layout.addWidget(preview_group)
        
        # 添加到堆叠式部件
        self.stacked_widget.addWidget(page)
    
    def create_step4_page(self):
        """创建步骤4：导出页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 20, 0, 0)
        layout.setSpacing(20)
        
        # 导出设置区域
        export_group = QGroupBox("导出PDF")
        export_layout = QVBoxLayout(export_group)
        
        # 导出摘要信息
        self.export_summary = QLabel("导出摘要信息")
        self.export_summary.setAlignment(Qt.AlignmentFlag.AlignCenter)
        export_layout.addWidget(self.export_summary)
        
        # 导出按钮
        export_btn_layout = QHBoxLayout()
        export_btn_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.export_btn = QPushButton("仅导出内容PDF")
        self.export_btn.setObjectName("accentButton")
        self.export_btn.setMinimumHeight(40)
        self.export_btn.clicked.connect(self.process_ppt)
        export_btn_layout.addWidget(self.export_btn)
        
        export_layout.addLayout(export_btn_layout)
        
        layout.addWidget(export_group)
        
        # 导出结果区域
        self.result_group = QGroupBox("导出结果")
        result_layout = QVBoxLayout(self.result_group)
        
        self.export_result = QLabel("点击上方按钮开始导出")
        self.export_result.setAlignment(Qt.AlignmentFlag.AlignCenter)
        result_layout.addWidget(self.export_result)
        
        layout.addWidget(self.result_group)
        
        # 添加AI索引选项
        self.ai_index_button = QPushButton("可选：添加AI索引 >")
        self.ai_index_button.clicked.connect(self.go_to_ai_step)
        self.ai_index_button.setVisible(False) # 导出成功后显示
        layout.addWidget(self.ai_index_button, 0, Qt.AlignmentFlag.AlignRight)
        
        # 添加到堆叠式部件
        self.stacked_widget.addWidget(page)
    
    def create_step5_page(self):
        """创建步骤5：AI索引生成页面"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 20, 0, 0)
        layout.setSpacing(15)

        # 1. 提示词生成区域
        prompt_group = QGroupBox("1. 生成并复制提示词")
        prompt_layout = QVBoxLayout(prompt_group)

        self.ai_prompt_text = QTextEdit()
        self.ai_prompt_text.setReadOnly(True)
        self.ai_prompt_text.setPlaceholderText("在这里生成给AI的提示词...")
        self.ai_prompt_text.setMinimumHeight(150)
        prompt_layout.addWidget(self.ai_prompt_text)

        copy_prompt_btn = QPushButton("复制提示词")
        copy_prompt_btn.clicked.connect(self.copy_ai_prompt)
        prompt_layout.addWidget(copy_prompt_btn, 0, Qt.AlignmentFlag.AlignRight)
        
        layout.addWidget(prompt_group)

        # 2. 粘贴Markdown区域
        markdown_group = QGroupBox("2. 粘贴AI返回的Markdown索引")
        markdown_layout = QVBoxLayout(markdown_group)

        self.ai_markdown_input = QTextEdit()
        self.ai_markdown_input.setPlaceholderText("请将AI生成的Markdown格式索引粘贴到此处。")
        self.ai_markdown_input.setMinimumHeight(200)
        markdown_layout.addWidget(self.ai_markdown_input)

        layout.addWidget(markdown_group)

        # 3. 生成最终PDF
        final_export_group = QGroupBox("3. 生成最终PDF")
        final_export_layout = QVBoxLayout(final_export_group)

        self.final_export_btn = QPushButton("合并生成带索引的PDF")
        self.final_export_btn.setObjectName("accentButton")
        self.final_export_btn.setMinimumHeight(40)
        self.final_export_btn.clicked.connect(self.generate_final_pdf_with_index)
        final_export_layout.addWidget(self.final_export_btn)

        self.final_export_result = QLabel()
        self.final_export_result.setAlignment(Qt.AlignmentFlag.AlignCenter)
        final_export_layout.addWidget(self.final_export_result)
        
        layout.addWidget(final_export_group)

        self.stacked_widget.addWidget(page)
    
    def show_welcome_screen(self):
        """显示欢迎界面"""
        # 设置初始步骤
        self.step_indicator.set_current_step(0)
        self.stacked_widget.setCurrentIndex(0)
        
        # 显示欢迎信息
        self.file_preview.setText(WELCOME_TEXT)
    
    def go_to_next_step(self):
        """前往下一步"""
        current_index = self.stacked_widget.currentIndex()
        
        # 验证当前步骤
        if current_index == 0 and not self.slide_images:
            QMessageBox.warning(self, "警告", "请先选择PPT文件！")
            return
        elif current_index == 2:
            # 如果从预览到导出，更新导出摘要信息
            layout_result = self.layout_calculator.calculate_layout(
                self.slide_images, self.layout_config
            )
            
            orientation_text = "横向" if self.layout_config["is_landscape"] else "纵向"
            summary = f"<p>将导出 <b>{len(self.slide_images)}</b> 张PPT幻灯片</p>"
            summary += f"<p>页面方向: <b>{orientation_text}A4</b></p>"
            summary += f"<p>布局: 每页 <b>{layout_result['rows']}</b> 行 × <b>{layout_result['columns']}</b> 列</p>"
            summary += f"<p>预计页数: <b>{layout_result['pages_needed']}</b> 页PDF</p>"
            summary += "<p>点击「仅导出内容PDF」按钮选择保存位置并开始导出</p>"
            
            self.export_summary.setText(summary)
            # 重置第四步状态
            self.export_result.setText("点击上方按钮开始导出")
            self.ai_index_button.setVisible(False)
            
        # 前往下一步
        if current_index < self.stacked_widget.count() - 2: # Stop before AI step
            self.stacked_widget.setCurrentIndex(current_index + 1)
            self.step_indicator.set_current_step(current_index + 1)
            
            # 如果是前往预览页面，立即刷新预览
            if current_index == 1:
                self.refresh_preview()
            
            # 更新按钮状态
            self.prev_btn.setEnabled(True)
            self.next_btn.setEnabled(current_index < self.stacked_widget.count() - 3)
    
    def go_to_prev_step(self):
        """返回上一步"""
        current_index = self.stacked_widget.currentIndex()
        
        if current_index > 0:
            self.stacked_widget.setCurrentIndex(current_index - 1)
            self.step_indicator.set_current_step(current_index - 1)
            
            # 更新按钮状态
            self.next_btn.setEnabled(True)
            self.prev_btn.setEnabled(current_index > 1)
    
    def select_ppt_file(self):
        """选择PPT文件并使用工作线程进行处理"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择PPT文件", "", "PowerPoint文件 (*.pptx *.ppt)"
        )
        
        if file_path:
            self.current_ppt_path = file_path
            self.file_info.setText(f"已选择: {os.path.basename(file_path)}")
            
            self.loading_overlay.set_text("正在转换PPT文件...")
            self.loading_overlay.show()

            # 关闭之前的图像以释放资源
            if self.slide_images:
                for img in self.slide_images:
                    try:
                        if hasattr(img, 'close'):
                            img.close()
                    except:
                        pass
                self.slide_images = []

            # 在工作线程中转换PPT
            self.worker = Worker(self.ppt_processor.convert_ppt_to_images, file_path)
            self.worker.finished.connect(self._on_ppt_conversion_finished)
            self.worker.error.connect(self._on_task_error)
            self.worker.start()

    def _on_ppt_conversion_finished(self, slide_images):
        """PPT转换完成后的回调"""
        self.loading_overlay.hide()
        self.slide_images = slide_images
        
        if self.slide_images:
            info = f"<p>成功加载 <b>{len(self.slide_images)}</b> 张PPT幻灯片</p>"
            info += "<p>点击「下一步」进行布局设置</p>"
            self.file_preview.setText(info)
            self.next_btn.setEnabled(True)
            if self.current_ppt_path:
                self.status_bar.showMessage(f"文件已加载: {os.path.basename(self.current_ppt_path)}")
        else:
            self.file_preview.setText(f"<p style='color:{COLORS['error']};'>文件处理失败</p><p>请选择有效的PPT文件或检查依赖项</p>")
            self.next_btn.setEnabled(False)
            self.status_bar.showMessage("文件处理失败")

    def _on_task_error(self, exception):
        """处理工作线程中的错误"""
        self.loading_overlay.hide()
        error_message = f"发生了一个错误: \n{str(exception)}"
        QMessageBox.critical(self, "操作失败", error_message)
        print(f"工作线程错误: {exception}")
        
        # 重置可能被禁用的按钮
        self.export_btn.setEnabled(True)
        self.final_export_btn.setEnabled(True)
        self.status_bar.showMessage("操作失败", 5000)

    def update_orientation(self):
        """更新页面方向设置"""
        is_landscape = self.landscape_radio.isChecked()
        if is_landscape != self.layout_config["is_landscape"]:
            self.layout_config["is_landscape"] = is_landscape
            orientation_text = "横向" if is_landscape else "纵向"
            self.status_bar.showMessage(f"页面方向已更改为{orientation_text}A4")
    
    def update_layout(self):
        """更新布局设置"""
        self.layout_config["columns"] = self.columns_spin.value()
        self.status_bar.showMessage(f"布局更新: 每行 {self.layout_config['columns']} 个PPT")
    
    def update_spacing(self):
        """更新间距设置"""
        self.layout_config["h_spacing"] = self.h_spacing_spin.value()
        self.layout_config["v_spacing"] = self.v_spacing_spin.value()
    
    def update_margins(self):
        """更新页边距设置"""
        self.layout_config["margin_left"] = self.margin_left_spin.value()
        self.layout_config["margin_right"] = self.margin_right_spin.value()
        self.layout_config["margin_top"] = self.margin_top_spin.value()
        self.layout_config["margin_bottom"] = self.margin_bottom_spin.value()
    
    def update_page_numbers(self):
        """更新页码设置"""
        self.layout_config["show_ppt_numbers"] = self.show_ppt_numbers_check.isChecked()
        self.layout_config["show_page_numbers"] = self.show_page_numbers_check.isChecked()
        page_numbers_text = []
        
        if self.layout_config["show_ppt_numbers"]:
            page_numbers_text.append("PPT定位页码")
            
        if self.layout_config["show_page_numbers"]:
            page_numbers_text.append("纸张页码")
            
        if page_numbers_text:
            self.status_bar.showMessage(f"已启用页码显示: {', '.join(page_numbers_text)}")
        else:
            self.status_bar.showMessage("已禁用所有页码显示")
    
    def refresh_preview(self):
        """刷新预览"""
        if not self.slide_images:
            return
        
        # 更新状态栏
        self.status_bar.showMessage("正在生成预览...")
        
        # 清除当前预览
        for i in reversed(range(self.preview_layout.count())):
            item = self.preview_layout.itemAt(i)
            if item.widget():
                item.widget().deleteLater()
            self.preview_layout.removeItem(item)
        
        # 计算布局
        layout_result = self.layout_calculator.calculate_layout(
            self.slide_images, self.layout_config
        )
        
        # 显示计算结果
        orientation_text = "横向" if self.layout_config["is_landscape"] else "纵向"
        result_text = f"<p>页面方向: <b>{orientation_text}A4</b></p>"
        result_text += f"<p>布局结果: 每页 <b>{layout_result['rows']}</b> 行 × <b>{layout_result['columns']}</b> 列</p>"
        result_text += f"<p>每个PPT尺寸: <b>{layout_result['item_width']:.1f}</b> × <b>{layout_result['item_height']:.1f}</b> mm</p>"
        result_text += f"<p>预计页数: <b>{layout_result['pages_needed']}</b> 页</p>"
        
        # 添加页码设置信息
        page_numbers_text = []
        if self.layout_config["show_ppt_numbers"]:
            page_numbers_text.append("PPT定位页码")
        if self.layout_config["show_page_numbers"]:
            page_numbers_text.append("纸张页码")
        
        if page_numbers_text:
            result_text += f"<p>页码显示: <b>{', '.join(page_numbers_text)}</b></p>"
        else:
            result_text += "<p>页码显示: <b>无</b></p>"
            
        self.preview_info.setText(result_text)
        
        # 创建预览
        preview_label = QLabel()
        preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # 根据方向设置画布大小
        if self.layout_config["is_landscape"]:
            pixmap = QPixmap(900, 650)  # 横向A4比例
        else:
            pixmap = QPixmap(650, 900)  # 纵向A4比例
            
        pixmap.fill(QColor(COLORS['surface']))
        
        painter = QPainter(pixmap)
        
        # 计算缩放因子，使A4页面适合画布大小
        if self.layout_config["is_landscape"]:
            scale_factor = pixmap.width() / 297  # A4横向宽度为297mm
        else:
            scale_factor = pixmap.width() / 210  # A4纵向宽度为210mm
        
        # 设置画笔
        pen = QPen(QColor(COLORS['divider']))
        pen.setWidth(1)
        painter.setPen(pen)
        
        # 根据方向获取页面尺寸
        page_width = 297 if self.layout_config["is_landscape"] else 210
        page_height = 210 if self.layout_config["is_landscape"] else 297
        
        # 绘制A4页面边框
        page_width_px = int(page_width * scale_factor)
        page_height_px = int(page_height * scale_factor)
        painter.fillRect(0, 0, page_width_px, page_height_px, QColor(COLORS['surface']))
        painter.setPen(QPen(QColor(COLORS['primary']), 2))
        painter.drawRect(0, 0, page_width_px, page_height_px)
        
        # 绘制PPT预览
        left_margin = self.layout_config["margin_left"] * scale_factor
        top_margin = self.layout_config["margin_top"] * scale_factor
        right_margin = self.layout_config["margin_right"] * scale_factor
        bottom_margin = self.layout_config["margin_bottom"] * scale_factor
        item_width = layout_result["item_width"] * scale_factor
        item_height = layout_result["item_height"] * scale_factor
        h_spacing = self.layout_config["h_spacing"] * scale_factor
        v_spacing = self.layout_config["v_spacing"] * scale_factor
        
        # 设置字体
        font = QFont("Arial", 9)
        painter.setFont(font)
        
        # 示例模拟几个PPT位置
        for row in range(layout_result["rows"]):
            for col in range(layout_result["columns"]):
                x = left_margin + col * (item_width + h_spacing)
                y = top_margin + row * (item_height + v_spacing)
                
                # 绘制PPT边框
                painter.setPen(QPen(QColor(COLORS['primary_dark']), 1))
                painter.fillRect(int(x), int(y), int(item_width), int(item_height), QColor(COLORS['background']))
                painter.drawRect(int(x), int(y), int(item_width), int(item_height))
                
                # 绘制编号
                painter.setPen(QColor(COLORS['text_primary']))
                item_num = row * layout_result["columns"] + col + 1
                if item_num <= len(self.slide_images):
                    painter.drawText(
                        QRectF(int(x), int(y), int(item_width), int(item_height)),
                        Qt.AlignmentFlag.AlignCenter,
                        f"PPT {item_num}"
                    )
                    
                    # 显示PPT定位页码
                    if self.layout_config["show_ppt_numbers"]:
                        painter.setFont(QFont("Arial", 8))
                        painter.drawText(
                            int(x), int(y + item_height + 12),
                            f"{item_num}"
                        )
                        painter.setFont(font)  # 恢复字体
        
        # 显示纸张页码
        if self.layout_config["show_page_numbers"]:
            painter.setFont(QFont("Arial", 10))
            page_number_text = "第 1 页 / 共 1 页"  # 预览中只显示第一页
            # 在右下角显示页码
            text_width = painter.fontMetrics().horizontalAdvance(page_number_text)
            painter.drawText(
                page_width_px - int(right_margin) - text_width - 10,
                page_height_px - int(bottom_margin),
                page_number_text
            )
        
        painter.end()
        
        # 显示预览图
        preview_label.setPixmap(pixmap)
        self.preview_layout.addWidget(preview_label)
        
        # 更新状态栏
        self.status_bar.showMessage("预览已生成")
        
        # 启用下一步按钮
        self.next_btn.setEnabled(True)
    
    def process_ppt(self):
        """处理PPT并导出PDF（异步）"""
        if not self.slide_images:
            return
        
        output_path, _ = QFileDialog.getSaveFileName(self, "保存PDF文件", "", "PDF文件 (*.pdf)")
        if not output_path:
            return

        self.content_pdf_path = output_path
        
        self.loading_overlay.set_text("正在生成内容PDF...")
        self.loading_overlay.show()
        self.export_btn.setEnabled(False)

        layout_result = self.layout_calculator.calculate_layout(self.slide_images, self.layout_config)
        
        self.worker = Worker(
            self.ppt_processor.generate_pdf, 
            self.slide_images, 
            self.content_pdf_path,
            layout_result, 
            self.layout_config
        )
        self.worker.finished.connect(self._on_content_pdf_generated)
        self.worker.error.connect(self._on_task_error)
        self.worker.start()

    def _on_content_pdf_generated(self, success):
        """内容PDF生成完成后的回调"""
        self.loading_overlay.hide()
        self.export_btn.setEnabled(True)

        if success:
            self.export_result.setText(f"<p style='color:{COLORS['success']};'><b>PDF导出成功!</b></p><p>文件保存在: {self.content_pdf_path}</p>")
            self.status_bar.showMessage(f"PDF已成功导出: {os.path.basename(self.content_pdf_path)}")
            QMessageBox.information(self, "导出成功", f"PDF已成功保存到:\n{self.content_pdf_path}")
            self.ai_index_button.setVisible(True)
        else:
            self.export_result.setText(f"<p style='color:{COLORS['error']};'><b>导出失败!</b></p><p>请检查文件权限和磁盘空间</p>")
            self.status_bar.showMessage("PDF导出失败")
            QMessageBox.critical(self, "导出失败", "生成PDF时出错，请检查日志。")

    def go_to_ai_step(self):
        """跳转到AI索引步骤"""
        self.step_indicator.set_current_step(4)
        self.stacked_widget.setCurrentIndex(4)
        
        # 生成并显示提示词
        self._generate_ai_prompt()

        # 更新按钮状态
        self.next_btn.setEnabled(False)
        self.prev_btn.setEnabled(True)
    
    def copy_ai_prompt(self):
        """复制AI提示词到剪贴板"""
        clipboard = QApplication.clipboard()
        clipboard.setText(self.ai_prompt_text.toPlainText())
        self.status_bar.showMessage("提示词已复制到剪贴板", 3000)
    
    def _generate_ai_prompt(self):
        """生成给AI的提示词"""
        if not self.slide_images:
            return

        layout_result = self.layout_calculator.calculate_layout(
            self.slide_images, self.layout_config
        )
        total_slides = len(self.slide_images)
        items_per_page = layout_result["rows"] * layout_result["columns"]

        prompt = f"""
我有一份包含 {total_slides} 张幻灯片的演示文稿。
我已经将它排版成一个PDF文档，每页包含 {items_per_page} 张幻灯片。共有 {layout_result['pages_needed']} 页。

幻灯片的编号方式如下：
- 每页幻灯片从左上角开始，从1开始编号，按行从左到右，从上到下依次编号。
- 您在最终的索引中，只需要引用这个编号即可。

例如：
- 索引的计算方式为 "(floor(PPT位置 / 每页幻灯片数)+1) - (PPT位置 % 每页幻灯片数)"
- 一个幻灯片在PPT中的位置为14，则生成的索引为： (floor(14 / {items_per_page})+1)-{14 % items_per_page}
- 如果每页包含2张幻灯片，一个幻灯片在PPT中的页号为25，则其索引为： (floor(25 / 2)+1)-{25 % 2} = 13-1

请您根据我稍后提供的所有幻灯片内容，为我生成一份详细的、树状结构的知识点索引目录。
索引需要采用Markdown格式，包含清晰的层级关系（例如使用#、##、-等）。
请确保索引覆盖所有关键知识点，并准确地将每个知识点定位到对应的幻灯片编号。
内容请具体到各个内容标题下的知识点，给出的更多的知识点内容及其索引。
"""
        self.ai_prompt_text.setText(prompt.strip())
    
    def generate_final_pdf_with_index(self):
        """生成包含AI索引的最终PDF（异步）"""
        markdown_text = self.ai_markdown_input.toPlainText()
        if not markdown_text.strip():
            QMessageBox.warning(self, "警告", "请输入AI生成的Markdown索引内容。")
            return

        if not hasattr(self, 'content_pdf_path') or not self.content_pdf_path:
            QMessageBox.critical(self, "错误", "未找到已导出的内容PDF，请先完成第四步。")
            return
            
        final_output_path, _ = QFileDialog.getSaveFileName(
            self, "保存带索引的PDF文件", os.path.dirname(self.content_pdf_path), "PDF文件 (*.pdf)"
        )
        if not final_output_path:
            return
        
        self.loading_overlay.set_text("正在生成索引并合并PDF...")
        self.loading_overlay.show()
        self.final_export_btn.setEnabled(False)
        
        self.worker = Worker(
            self.ppt_processor.generate_pdf_with_index,
            markdown_text,
            self.content_pdf_path,
            final_output_path
        )
        self.worker.finished.connect(lambda success: self._on_final_pdf_generated(success, final_output_path))
        self.worker.error.connect(self._on_task_error)
        self.worker.start()

    def _on_final_pdf_generated(self, success, final_output_path):
        """最终PDF生成完成后的回调"""
        self.loading_overlay.hide()
        self.final_export_btn.setEnabled(True)

        if success:
            self.final_export_result.setText(f"<p style='color:{COLORS['success']};'><b>带索引的PDF导出成功!</b></p><p>文件保存在: {final_output_path}</p>")
            QMessageBox.information(self, "成功", f"带索引的最终PDF已保存到:\n{final_output_path}")
        else:
            self.final_export_result.setText(f"<p style='color:{COLORS['error']};'><b>最终PDF生成失败!</b></p><p>请检查日志获取详细信息。</p>")
            QMessageBox.critical(self, "失败", "生成最终PDF时出错，请检查日志。")
    
    def closeEvent(self, event):
        """程序关闭时清理临时文件"""
        try:
            # 清理PPT处理器的临时文件
            if hasattr(self, 'ppt_processor') and self.ppt_processor:
                self.ppt_processor.cleanup_temp_files()
                self.status_bar.showMessage("已清理所有临时文件")
            
            # 关闭所有可能仍然打开的图像
            if hasattr(self, 'slide_images') and self.slide_images:
                for img in self.slide_images:
                    try:
                        if hasattr(img, 'close'):
                            img.close()
                    except:
                        pass
            
            # 继续正常关闭
            event.accept()
        except Exception as e:
            print(f"关闭时发生错误: {e}")
            event.accept() 