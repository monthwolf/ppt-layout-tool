import os
import sys
from pathlib import Path
import requests
from packaging import version
from PyQt6.QtWidgets import (QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, 
                           QPushButton, QSpinBox, QLabel, QFileDialog, 
                           QScrollArea, QGroupBox, QDoubleSpinBox, QMessageBox,
                           QSizePolicy, QFrame, QGridLayout,
                           QStatusBar, QStackedWidget, QRadioButton, QButtonGroup,
                           QCheckBox, QTextEdit, QApplication)
from PyQt6.QtCore import Qt, QRectF, QPropertyAnimation, QEasingCurve, QParallelAnimationGroup, QRect, QSettings, QUrl, QTimer
from PyQt6.QtGui import QPixmap, QPainter, QPen, QColor, QFont, QIcon, QAction, QDesktopServices
from PyQt6.QtSvg import QSvgRenderer

from pptx import Presentation
from PIL import Image
import io

from src.utils.ppt_processor import PPTProcessor
from src.utils.layout_calculator import LayoutCalculator
from src.ui.styles import STYLESHEET, COLORS, WELCOME_TEXT, STEPS_GUIDE
from src.ui.loading_overlay import LoadingOverlay
from src.ui.worker import Worker
from src.ui.spinner_widget import SpinnerWidget

# 从环境变量读取版本号，如果未设置则为开发版
CURRENT_VERSION = os.environ.get("APP_VERSION", "dev")
GITHUB_REPO = "monthwolf/ppt-layout-tool" # 请替换为您的GitHub仓库

def get_resource_path(relative_path):
    """一个健壮的函数，用于在开发和打包环境中都能找到资源文件。"""
    try:
        # PyInstaller 在运行时会创建一个临时文件夹，并将其路径存储在 _MEIPASS 中
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        # 在开发环境中，我们从当前文件向上回溯两层以找到项目根目录
        base_path = Path(__file__).resolve().parents[2]
    return str(Path(base_path) / relative_path)

class AnimatedStackedWidget(QStackedWidget):
    """一个支持平滑滑动动画的QStackedWidget"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.m_speed = 350
        self.m_animation_type = QEasingCurve.Type.InOutCubic
        self.m_now = 0
        self.m_next = 0
        self.m_active = False

    def slide_in_next(self):
        now_widget = self.widget(self.m_now)
        next_widget = self.widget(self.m_next)
        
        anim_group = QParallelAnimationGroup(self)

        anim_now = QPropertyAnimation(now_widget, b"geometry") # type: ignore
        anim_now.setDuration(self.m_speed)
        anim_now.setEasingCurve(self.m_animation_type)
        anim_now.setStartValue(QRect(0, 0, self.width(), self.height()))
        anim_now.setEndValue(QRect(-self.width(), 0, self.width(), self.height()))
        anim_group.addAnimation(anim_now)

        anim_next = QPropertyAnimation(next_widget, b"geometry") # type: ignore
        anim_next.setDuration(self.m_speed)
        anim_next.setEasingCurve(self.m_animation_type)
        anim_next.setStartValue(QRect(self.width(), 0, self.width(), self.height()))
        anim_next.setEndValue(QRect(0, 0, self.width(), self.height()))
        anim_group.addAnimation(anim_next)
        
        anim_group.finished.connect(self.animation_done)
        self.m_active = True
        anim_group.start()

    def slide_in_prev(self):
        now_widget = self.widget(self.m_now)
        next_widget = self.widget(self.m_next)

        anim_group = QParallelAnimationGroup(self)

        anim_now = QPropertyAnimation(now_widget, b"geometry") # type: ignore
        anim_now.setDuration(self.m_speed)
        anim_now.setEasingCurve(self.m_animation_type)
        anim_now.setStartValue(QRect(0, 0, self.width(), self.height()))
        anim_now.setEndValue(QRect(self.width(), 0, self.width(), self.height()))
        anim_group.addAnimation(anim_now)

        anim_next = QPropertyAnimation(next_widget, b"geometry") # type: ignore
        anim_next.setDuration(self.m_speed)
        anim_next.setEasingCurve(self.m_animation_type)
        anim_next.setStartValue(QRect(-self.width(), 0, self.width(), self.height()))
        anim_next.setEndValue(QRect(0, 0, self.width(), self.height()))
        anim_group.addAnimation(anim_next)

        anim_group.finished.connect(self.animation_done)
        self.m_active = True
        anim_group.start()

    def setCurrentIndex(self, index):
        if self.m_active or index == self.currentIndex():
            return
        
        self.m_next = index
        next_widget = self.widget(self.m_next)
        next_widget.setGeometry(0, 0, self.width(), self.height())

        if index > self.m_now:
            next_widget.move(self.width(), 0)
        elif index < self.m_now:
            next_widget.move(-self.width(), 0)
        
        next_widget.show()
        if index > self.m_now:
            self.slide_in_next()
        else:
            self.slide_in_prev()

    def animation_done(self):
        self.widget(self.m_now).hide()
        self.setCurrentWidget(self.widget(self.m_next))
        self.m_active = False
        self.m_now = self.m_next

class StepIndicator(QFrame):
    """一个现代化的、带动画效果的步骤指示器"""
    def __init__(self, steps, parent=None):
        super().__init__(parent)
        self.steps = steps
        self.current_step = 0
        
        layout = QHBoxLayout(self)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 10, 0, 10)
        
        self.step_widgets = []
        for i, step_text in enumerate(steps):
            step_container = QWidget()
            step_layout = QVBoxLayout(step_container)
            step_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            step_layout.setContentsMargins(5, 0, 5, 0)

            icon_stack = QStackedWidget()
            icon_stack.setFixedSize(32, 32)
            
            number_label = QLabel(str(i + 1))
            number_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            number_label.setStyleSheet(f"background-color: {COLORS['divider']}; border-radius: 16px; color: {COLORS['text_secondary']}; font-weight: normal;")
            icon_stack.addWidget(number_label)

            spinner = SpinnerWidget()
            icon_stack.addWidget(spinner)

            check_label = QLabel()
            check_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            check_label.setPixmap(self._create_check_pixmap())
            icon_stack.addWidget(check_label)
            
            text_label = QLabel(step_text)
            
            step_layout.addWidget(icon_stack, 0, Qt.AlignmentFlag.AlignCenter)
            step_layout.addWidget(text_label, 0, Qt.AlignmentFlag.AlignCenter)
            
            layout.addWidget(step_container)
            self.step_widgets.append({'icons': icon_stack, 'text': text_label, 'spinner': spinner})
            
            if i < len(steps) - 1:
                line = QFrame()
                line.setFrameShape(QFrame.Shape.HLine)
                line.setFixedHeight(2)
                line.setStyleSheet(f"background-color: {COLORS['divider']};")
                line.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
                layout.addWidget(line)

    def _create_check_pixmap(self):
        pixmap = QPixmap(32, 32)
        pixmap.fill(Qt.GlobalColor.transparent)
        p = QPainter(pixmap)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setBrush(QColor(COLORS['success']))
        p.setPen(Qt.GlobalColor.transparent)
        p.drawEllipse(0, 0, 32, 32)
        
        renderer = QSvgRenderer(get_resource_path("resources/check.svg"))
        renderer.render(p, QRectF(6, 6, 20, 20))
        p.end()
        return pixmap

    def set_current_step(self, step):
        if 0 <= step < len(self.steps):
            for i, widget_group in enumerate(self.step_widgets):
                widget_group['spinner'].stop()

            for i, widget_group in enumerate(self.step_widgets):
                icons = widget_group['icons']
                text = widget_group['text']
                if i < step:
                    icons.setCurrentIndex(2)
                    text.setStyleSheet(f"color: {COLORS['text_secondary']}; font-weight: normal;")
                elif i == step:
                    icons.setCurrentIndex(1)
                    widget_group['spinner'].start()
                    text.setStyleSheet(f"color: {COLORS['primary']}; font-weight: bold;")
                else:
                    icons.setCurrentIndex(0)
                    text.setStyleSheet(f"color: {COLORS['text_secondary']}; font-weight: normal;")
            self.current_step = step

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"PPT 布局工具 v{CURRENT_VERSION}")
        self.setMinimumSize(1100, 800)
        self.setWindowIcon(QIcon(get_resource_path("resources/app_icon.svg")))

        self.ppt_processor = PPTProcessor()
        self.layout_calculator = LayoutCalculator()
        
        self.content_pdf_path = None
        self.current_ppt_path = None
        self.slide_images = []
        self.layout_config = {
            "columns": 2, "page_width": 210, "page_height": 297,
            "margin_left": 10, "margin_top": 10, "margin_right": 10, "margin_bottom": 10,
            "h_spacing": 5, "v_spacing": 5, "is_landscape": True,
            "show_ppt_numbers": True, "show_page_numbers": True,
        }
        
        self.setStyleSheet(STYLESHEET)
        
        self.init_ui()
        self.init_loading_overlay()
        self.init_menu()
        
        self.show_welcome_screen()
        
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("准备就绪")

        # 启动时检查
        QTimer.singleShot(100, self.initial_checks)

    def initial_checks(self):
        """执行启动时的检查，如更新检查和首次启动欢迎"""
        # 启动时静默检查更新
        self.check_for_updates(silent=True)

    def check_first_launch(self):
        """检查是否首次启动，如果是则显示关于对话框"""
        settings = QSettings("monthwolf", "PPTLayoutTool")
        if settings.value("showAboutOnLaunch", True, type=bool):
            self.show_about_dialog()
            
    def show_about_dialog(self):
        """显示关于对话框"""
        # 创建一个自定义的QMessageBox
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("关于 PPT 布局工具")
        
        about_text = f"""
            <h2 style="font-size: 18px; color: #333;">PPT 布局工具 v{CURRENT_VERSION}</h2>
            <p style="font-size: 13px; color: #555;">一款现代、高效的PPT排版与索引生成工具。</p>
            <p style="font-size: 13px; color: #555;">作者：<b>monthwolf</b></p>
            <p style="font-size: 13px; color: #555;">
                <a href="https://github.com/{GITHUB_REPO}" style="color: #0078D7;">访问 GitHub 仓库</a> | 
                <a href="https://github.com/{GITHUB_REPO}/releases" style="color: #0078D7;">查看所有版本</a>
            </p>
        """
        msg_box.setText(about_text)
        msg_box.setIcon(QMessageBox.Icon.Information)

        # 添加"不再显示"复选框
        cb = QCheckBox("启动时不再显示此对话框")
        settings = QSettings("monthwolf", "PPTLayoutTool")
        show_on_launch = settings.value("showAboutOnLaunch", True, type=bool)
        cb.setChecked(not show_on_launch)
        msg_box.setCheckBox(cb)
        
        msg_box.exec()

        # 根据复选框状态更新设置
        settings.setValue("showAboutOnLaunch", not cb.isChecked())
            
    def init_menu(self):
        menu_bar = self.menuBar()
        help_menu = menu_bar.addMenu("帮助")

        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)
        
        update_action = QAction("检查更新", self)
        update_action.triggered.connect(lambda: self.check_for_updates(silent=False))
        help_menu.addAction(update_action)

    def init_loading_overlay(self):
        self.loading_overlay = LoadingOverlay(self)
    
    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'loading_overlay'):
            self.loading_overlay.resize(event.size())

    def init_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        
        self.step_indicator = StepIndicator(STEPS_GUIDE)
        main_layout.addWidget(self.step_indicator)
        
        self.stacked_widget = AnimatedStackedWidget()
        
        self.create_step1_page()
        self.create_step2_page()
        self.create_step3_page()
        self.create_step4_page()
        self.create_step5_page()
        
        main_layout.addWidget(self.stacked_widget)
        
        nav_layout = QHBoxLayout()
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        nav_layout.addWidget(spacer)
        
        self.prev_btn = QPushButton("< 上一步")
        self.prev_btn.clicked.connect(self.go_to_prev_step)
        self.prev_btn.setEnabled(False)
        nav_layout.addWidget(self.prev_btn)
        
        self.next_btn = QPushButton("下一步 >")
        self.next_btn.setObjectName("accentButton")
        self.next_btn.clicked.connect(self.go_to_next_step)
        self.next_btn.setEnabled(False)
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
        
        # 导出摘要区域
        summary_group = QGroupBox("导出摘要")
        summary_layout = QVBoxLayout(summary_group)
        
        self.export_summary = QLabel("准备导出...")
        self.export_summary.setAlignment(Qt.AlignmentFlag.AlignCenter)
        summary_layout.addWidget(self.export_summary)
        
        layout.addWidget(summary_group)
        
        # 导出区域
        export_group = QGroupBox("导出PDF")
        export_layout = QVBoxLayout(export_group)
        
        export_hint = QLabel("点击下方按钮导出排版后的PDF文件")
        export_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        export_layout.addWidget(export_hint)
        
        # 导出按钮
        export_btn_layout = QHBoxLayout()
        export_btn_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.export_btn = QPushButton("仅导出内容PDF")
        self.export_btn.setMinimumSize(180, 40)
        self.export_btn.clicked.connect(self.process_ppt)
        export_btn_layout.addWidget(self.export_btn)
        
        export_layout.addLayout(export_btn_layout)
        
        # 导出结果
        self.export_result = QLabel()
        self.export_result.setAlignment(Qt.AlignmentFlag.AlignCenter)
        export_layout.addWidget(self.export_result)
        
        # AI索引按钮
        self.ai_index_button = QPushButton("可选：添加AI索引 >")
        self.ai_index_button.setObjectName("accentButton")
        self.ai_index_button.setMinimumSize(180, 40)
        self.ai_index_button.setVisible(False)  # 初始不可见
        self.ai_index_button.clicked.connect(self.go_to_ai_step)  # 确保连接到正确的方法
        export_layout.addWidget(self.ai_index_button, 0, Qt.AlignmentFlag.AlignCenter)
        
        layout.addWidget(export_group)
        
        self.stacked_widget.addWidget(page)
    
    def create_step5_page(self):
        """创建步骤5：AI索引页面（可选）"""
        page = QWidget()
        
        # 创建主滚动区域
        scroll_area = QScrollArea(page)
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        
        # 创建一个容器widget来放置所有的内容
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(0, 20, 0, 20)  # 添加底部边距，确保滚动时内容不被截断
        layout.setSpacing(20)
        
        # AI提示词区域
        prompt_group = QGroupBox("AI提示词")
        prompt_layout = QVBoxLayout(prompt_group)
        
        prompt_hint = QLabel("以下是生成的AI提示词，您可以将其复制到ChatGPT或其他AI工具中以获取更好的索引内容。")
        prompt_hint.setWordWrap(True)
        prompt_layout.addWidget(prompt_hint)
        
        copy_btn_layout = QHBoxLayout()
        copy_btn_layout.setAlignment(Qt.AlignmentFlag.AlignRight)
        
        copy_prompt_btn = QPushButton("复制提示词")
        copy_prompt_btn.clicked.connect(self.copy_ai_prompt)
        copy_btn_layout.addWidget(copy_prompt_btn)
        
        prompt_layout.addLayout(copy_btn_layout)
        
        self.ai_prompt_text = QTextEdit()
        self.ai_prompt_text.setReadOnly(True)
        self.ai_prompt_text.setMinimumHeight(100)
        prompt_layout.addWidget(self.ai_prompt_text)
        
        layout.addWidget(prompt_group)
        
        # AI生成的索引输入区域
        ai_group = QGroupBox("AI生成的Markdown索引")
        ai_layout = QVBoxLayout(ai_group)
        
        ai_hint = QLabel("请将AI生成的Markdown格式索引粘贴在此处：")
        ai_layout.addWidget(ai_hint)
        
        self.ai_markdown_input = QTextEdit()
        self.ai_markdown_input.setMinimumHeight(200)
        self.ai_markdown_input.setAcceptRichText(False)
        self.ai_markdown_input.textChanged.connect(self._update_export_button_state)
        ai_layout.addWidget(self.ai_markdown_input)
        
        layout.addWidget(ai_group)
        
        # 最终导出区域
        final_export_group = QGroupBox("导出带索引的PDF")
        final_export_layout = QVBoxLayout(final_export_group)
        
        self.final_export_btn = QPushButton("生成带索引的最终PDF")
        self.final_export_btn.setObjectName("accentButton")
        self.final_export_btn.setMinimumHeight(40)
        self.final_export_btn.clicked.connect(self.export_final_pdf)
        final_export_layout.addWidget(self.final_export_btn)
        
        self.final_export_result = QLabel()
        self.final_export_result.setAlignment(Qt.AlignmentFlag.AlignCenter)
        final_export_layout.addWidget(self.final_export_result)
        
        layout.addWidget(final_export_group)
        
        # 设置滚动区域的内容
        scroll_area.setWidget(scroll_content)
        
        # 设置主页面布局
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.addWidget(scroll_area)

        self.stacked_widget.addWidget(page)
    
    def _update_export_button_state(self):
        """根据Markdown输入框的内容更新导出按钮状态"""
        has_content = bool(self.ai_markdown_input.toPlainText().strip())
        self.final_export_btn.setEnabled(has_content)
    
    def show_welcome_screen(self):
        """显示欢迎界面"""
        # 设置初始步骤
        self.step_indicator.set_current_step(0)
        self.stacked_widget.setCurrentIndex(0)
        
        # 显示欢迎信息
        self.file_preview.setText(WELCOME_TEXT)
    
    def go_to_next_step(self):
        current_index = self.stacked_widget.currentIndex()
        
        if current_index == 0 and not self.slide_images:
            QMessageBox.warning(self, "警告", "请先选择PPT文件！")
            return
        elif current_index == 3:
            # 第四步结束后不能直接进入第五步，需要通过 AI索引按钮
            return
            
        if current_index < self.stacked_widget.count() - 1:
            if current_index == 2:
                self._update_export_summary()
            
            self.step_indicator.set_current_step(current_index + 1)
            self.stacked_widget.setCurrentIndex(current_index + 1)
            
            # 在进入第三步时自动生成预览
            next_index = current_index + 1
            if next_index == 2:  # 第三步索引为2
                # 显示加载覆盖层
                self.loading_overlay.set_text("正在生成布局预览...")
                self.loading_overlay.show()
                
                # 使用延时确保UI已经更新
                QTimer.singleShot(100, self._generate_preview_with_loading)
            
            self.prev_btn.setEnabled(True)
            
            # 更新下一步按钮状态
            # 当到达第四步时，禁用下一步按钮（因为第五步是可选的）
            is_last_regular_step = current_index >= self.stacked_widget.count() - 3  # 倒数第二个常规步骤
            self.next_btn.setEnabled(not is_last_regular_step)

    def _generate_preview_with_loading(self):
        """带加载覆盖层的预览生成"""
        try:
            # 刷新预览
            self.refresh_preview()
        finally:
            # 无论成功与否，都隐藏加载覆盖层
            self.loading_overlay.hide()

    def go_to_prev_step(self):
        current_index = self.stacked_widget.currentIndex()
        if current_index > 0:
            self.step_indicator.set_current_step(current_index - 1)
            self.stacked_widget.setCurrentIndex(current_index - 1)
            
            # 当从第五步返回时，启用下一步按钮
            is_from_optional_step = current_index == self.stacked_widget.count() - 1
            if is_from_optional_step:
                self.next_btn.setEnabled(False)  # 因为回到第四步，第四步的下一步按钮不可用
            else:
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
            
            # 显示进度条
            self.loading_overlay.set_progress(0, 100, "准备处理PPT文件...")

            # 关闭之前的图像以释放资源
            if self.slide_images:
                for img in self.slide_images:
                    try:
                        if hasattr(img, 'close'):
                            img.close()
                    except:
                        pass
                self.slide_images = []

            # 定义进度回调函数
            def progress_callback(current, total, message):
                # 在UI线程中更新进度条
                self.loading_overlay.set_progress(current, total, message)

            # 在工作线程中转换PPT，并传递进度回调
            self.worker = Worker(
                self.ppt_processor.convert_ppt_to_images, 
                file_path,
                progress_callback=self._update_progress
            )
            self.worker.finished.connect(self._on_ppt_conversion_finished)
            self.worker.error.connect(self._on_task_error)
            
            # 连接进度信号
            self.worker.progress.connect(self._update_progress)
            
            self.worker.start()

    def _update_progress(self, current, total, message):
        """更新加载覆盖层的进度"""
        self.loading_overlay.set_progress(current, total, message)

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
        """任务错误处理"""
        self.loading_overlay.hide()
        self.export_btn.setEnabled(True)
        self.final_export_btn.setEnabled(True)  # 确保最终导出按钮也会被重新启用
        
        error_msg = str(exception)
        print(f"任务执行出错: {error_msg}")
        
        # 根据当前所处的步骤更新UI反馈
        current_index = self.stacked_widget.currentIndex()
        
        if current_index == 3:  # 导出页面
            self.export_result.setText(f"<p style='color:{COLORS['error']};'><b>PDF导出失败!</b></p><p>错误: {error_msg}</p>")
        elif current_index == 4:  # AI索引页面
            self.final_export_result.setText(f"<p style='color:{COLORS['error']};'><b>PDF生成失败!</b></p><p>错误: {error_msg}</p>")
        
        QMessageBox.critical(self, "错误", f"操作失败：{error_msg}")

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
        self.loading_overlay.set_progress(0, 100, "准备生成PDF...")
        self.export_btn.setEnabled(False)

        layout_result = self.layout_calculator.calculate_layout(self.slide_images, self.layout_config)
        
        # 定义进度回调函数
        def progress_callback(current, total, message):
            # 在UI线程中更新进度条
            self.loading_overlay.set_progress(current, total, message)
        
        self.worker = Worker(
            self.ppt_processor.generate_pdf, 
            self.slide_images, 
            self.content_pdf_path,
            layout_result, 
            self.layout_config,
            progress_callback=self._update_progress
        )
        self.worker.finished.connect(lambda success: self._on_content_pdf_generated(success, self.content_pdf_path))
        self.worker.error.connect(self._on_task_error)
        self.worker.progress.connect(self._update_progress)
        self.worker.start()

    def _on_content_pdf_generated(self, success, output_path):
        """内容PDF生成完成后的回调"""
        self.loading_overlay.hide()
        self.export_btn.setEnabled(True)

        if success:
            self.export_result.setText(f"<p style='color:{COLORS['success']};'><b>PDF导出成功!</b></p><p>文件保存在: {output_path}</p>")
            QMessageBox.information(self, "成功", f"PDF已保存到:\n{output_path}")
            
            # 显示AI索引按钮
            self.ai_index_button.setVisible(True)
        else:
            self.export_result.setText(f"<p style='color:{COLORS['error']};'><b>PDF导出失败!</b></p><p>请检查日志获取详细信息。</p>")
            QMessageBox.critical(self, "失败", "生成PDF时出错，请检查日志。")

    def go_to_ai_step(self):
        """跳转到AI索引步骤（第五步）"""
        if self.content_pdf_path and os.path.exists(self.content_pdf_path):
            self.step_indicator.set_current_step(4)  # 第五步的索引为4
            self.stacked_widget.setCurrentIndex(4)
            
            # 生成并显示提示词
            self._generate_ai_prompt()
            
            # 重置Markdown输入区域
            self.ai_markdown_input.clear()
            self.final_export_btn.setEnabled(False)
            
            # 更新按钮状态
            self.prev_btn.setEnabled(True)
            self.next_btn.setEnabled(False)
        else:
            QMessageBox.warning(self, "警告", "请先导出内容PDF！")
            return
    
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
    
    def export_final_pdf(self):
        """生成最终PDF（带索引）"""
        if not self.content_pdf_path or not os.path.exists(self.content_pdf_path):
            QMessageBox.warning(self, "警告", "请先生成内容PDF！")
            return
            
        markdown_text = self.ai_markdown_input.toPlainText().strip()
        if not markdown_text:
            QMessageBox.warning(self, "警告", "请输入AI生成的Markdown索引内容！")
            return
            
        final_output_path, _ = QFileDialog.getSaveFileName(self, "保存最终PDF文件", "", "PDF文件 (*.pdf)")
        if not final_output_path:
            return
            
        self.loading_overlay.set_text("正在生成带索引的PDF...")
        self.loading_overlay.show()
        self.loading_overlay.set_progress(0, 100, "准备生成索引PDF...")
        self.final_export_btn.setEnabled(False)
        
        # 定义进度回调函数
        def progress_callback(current, total, message):
            # 在UI线程中更新进度条
            self.loading_overlay.set_progress(current, total, message)
        
        # 使用异步线程生成PDF
        self.worker = Worker(
            self.ppt_processor.generate_pdf_with_index, 
            markdown_text,
            self.content_pdf_path,
            final_output_path,
            progress_callback=self._update_progress
        )
        self.worker.finished.connect(lambda success: self._on_final_pdf_generated(success, final_output_path))
        self.worker.error.connect(self._on_task_error)
        self.worker.progress.connect(self._update_progress)
        self.worker.start()

    def _on_final_pdf_generated(self, success, final_output_path):
        """最终PDF生成完成后的回调"""
        self.loading_overlay.hide()
        self.final_export_btn.setEnabled(True)

        if success and final_output_path:
            self.final_export_result.setText(f"<p style='color:{COLORS['success']};'><b>带索引的PDF导出成功!</b></p><p>文件保存在: {final_output_path}</p>")
            QMessageBox.information(self, "成功", f"带索引的最终PDF已保存到:\n{final_output_path}")
        else:
            self.final_export_result.setText(f"<p style='color:{COLORS['error']};'><b>最终PDF生成失败!</b></p><p>请检查日志获取详细信息。</p>")
            QMessageBox.critical(self, "失败", "生成最终PDF时出错，请检查日志。")
    
    def _update_export_summary(self):
        layout_result = self.layout_calculator.calculate_layout(self.slide_images, self.layout_config)
        orientation_text = "横向" if self.layout_config["is_landscape"] else "纵向"
        summary = f"<p>将导出 <b>{len(self.slide_images)}</b> 张PPT幻灯片</p>"
        summary += f"<p>页面方向: <b>{orientation_text}A4</b></p>"
        summary += f"<p>布局: 每页 <b>{layout_result['rows']}</b> 行 × <b>{layout_result['columns']}</b> 列</p>"
        summary += f"<p>预计页数: <b>{layout_result['pages_needed']}</b> 页PDF</p>"
        self.export_summary.setText(summary)
        
        # 隐藏AI索引按钮 - 仅在成功导出PDF后显示
        self.ai_index_button.setVisible(False)
    
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

    def check_for_updates(self, silent=False):
        """在工作线程中检查GitHub上的新版本"""
        if GITHUB_REPO == "YOUR_USERNAME/YOUR_REPOSITORY" or CURRENT_VERSION == "dev":
            if not silent:
                if CURRENT_VERSION == "dev":
                    QMessageBox.information(self, "开发版", "您正在运行开发版本，已跳过更新检查。")
                else:
                    QMessageBox.warning(self, "未配置", "GitHub仓库地址未配置，无法检查更新。")
            else:
                self.check_first_launch()
            return

        self.status_bar.showMessage("正在检查更新...")
        # 创建一个Worker来执行网络请求
        self.update_worker = Worker(self._get_latest_release_info)
        self.update_worker.finished.connect(lambda release_info: self._on_update_check_finished(release_info, silent))
        self.update_worker.error.connect(lambda e: self._on_update_check_error(e, silent))
        self.update_worker.start()

    def _get_latest_release_info(self):
        """从GitHub API获取最新发布信息"""
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        response = requests.get(api_url, timeout=10)
        response.raise_for_status()
        return response.json()

    def _on_update_check_finished(self, release_info, silent):
        """更新检查完成后的回调"""
        # 如果是开发版，不进行版本比较
        if CURRENT_VERSION == "dev":
            if not silent:
                QMessageBox.information(self, "开发版", "您正在运行开发版本，已跳过更新检查。")
            else:
                self.check_first_launch()
            return

        latest_version_str = release_info.get("tag_name", "v0.0.0").lstrip('v')
        
        try:
            current_v = version.parse(CURRENT_VERSION)
            latest_v = version.parse(latest_version_str)

            if latest_v > current_v:
                self.show_update_dialog(release_info)
            else:
                if not silent:
                    QMessageBox.information(self, "检查更新", "您当前使用的已是最新版本。")
                else:
                    self.check_first_launch()
        except version.InvalidVersion:
            if not silent:
                QMessageBox.warning(self, "版本错误", "无法解析版本号，已跳过更新检查。")
            else:
                self.check_first_launch()

    def _on_update_check_error(self, exception, silent):
        """更新检查失败时的回调"""
        print(f"检查更新失败: {exception}")
        if not silent:
            QMessageBox.warning(self, "检查更新", "无法连接到GitHub检查更新，请检查您的网络连接。")
        else:
            # 如果检查失败，同样继续显示关于对话框
            self.check_first_launch()
            
    def show_update_dialog(self, release_info):
        """显示更新提示对话框"""
        latest_version = release_info.get("tag_name", "v0.0.0")
        release_notes = release_info.get("body", "无版本说明。").replace('\n', '<br>')
        release_url = release_info.get("html_url", f"https://github.com/{GITHUB_REPO}/releases")

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("发现新版本！")
        msg_box.setIcon(QMessageBox.Icon.Information)
        
        update_text = f"""
            <h3 style="font-size: 16px;">发现新版本: {latest_version}</h3>
            <p style="font-size: 13px;">建议您更新到最新版本以获得最佳体验。</p>
            <p style="font-size: 14px; font-weight: bold;">更新内容:</p>
            <div style="font-size: 12px; background-color: #f0f0f0; border-radius: 5px; padding: 10px;">
                {release_notes}
            </div>
        """
        msg_box.setText(update_text)
        
        update_btn = msg_box.addButton("立即更新", QMessageBox.ButtonRole.ActionRole)
        msg_box.addButton("稍后提醒", QMessageBox.ButtonRole.RejectRole)
        
        msg_box.exec()
        
        if msg_box.clickedButton() == update_btn:
            QDesktopServices.openUrl(QUrl(release_url)) 