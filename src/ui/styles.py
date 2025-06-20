"""
定义应用程序的现代UI样式
"""

# 应用程序的主题颜色
COLORS = {
    'primary': '#2979FF',    # 主题色，蓝色
    'primary_light': '#75A7FF',
    'primary_dark': '#004ECB',
    'secondary': '#651FFF',  # 辅助色，紫色
    'background': '#FAFAFA',
    'surface': '#FFFFFF',
    'error': '#B00020',
    'success': '#4CAF50',
    'text_primary': '#212121',
    'text_secondary': '#757575',
    'divider': '#E0E0E0',
}

# 应用程序样式表
STYLESHEET = f"""
/* 全局样式 */
QWidget {{
    font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif;
    color: {COLORS['text_primary']};
    background-color: {COLORS['background']};
}}

/* 主窗口 */
QMainWindow {{
    background-color: {COLORS['background']};
}}

/* 按钮样式 */
QPushButton {{
    background-color: {COLORS['primary']};
    color: white;
    border: none;
    border-radius: 4px;
    padding: 8px 16px;
    font-weight: 500;
    min-width: 88px;
}}

QPushButton:hover {{
    background-color: {COLORS['primary_light']};
}}

QPushButton:pressed {{
    background-color: {COLORS['primary_dark']};
}}

QPushButton:disabled {{
    background-color: #BDBDBD;
    color: #757575;
}}

/* 强调按钮 */
QPushButton#accentButton {{
    background-color: {COLORS['secondary']};
    font-weight: bold;
}}

QPushButton#accentButton:hover {{
    background-color: #7C4DFF;
}}

/* 分组框 */
QGroupBox {{
    border: 1px solid {COLORS['divider']};
    border-radius: 4px;
    margin-top: 1.1em;
    padding-top: 0.5em;
    font-weight: 500;
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top center;
    padding: 0 10px;
    background-color: {COLORS['background']};
}}

/* 标签 */
QLabel {{
    color: {COLORS['text_primary']};
}}

QLabel#infoLabel {{
    color: {COLORS['text_secondary']};
    font-style: italic;
}}

QLabel#errorLabel {{
    color: {COLORS['error']};
}}

QLabel#successLabel {{
    color: {COLORS['success']};
}}

/* 输入框 */
QLineEdit, QSpinBox, QDoubleSpinBox {{
    border: 1px solid {COLORS['divider']};
    border-radius: 4px;
    padding: 4px;
    background-color: {COLORS['surface']};
}}

QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus {{
    border: 2px solid {COLORS['primary']};
}}

/* 滚动区域 */
QScrollArea {{
    border: 1px solid {COLORS['divider']};
    border-radius: 4px;
    background-color: {COLORS['surface']};
}}

/* 工具提示 */
QToolTip {{
    background-color: #424242;
    color: white;
    border: none;
    border-radius: 2px;
    padding: 4px;
}}

/* 状态栏 */
QStatusBar {{
    background-color: {COLORS['surface']};
    color: {COLORS['text_secondary']};
}}

/* 进度条 */
QProgressBar {{
    border: none;
    border-radius: 4px;
    background-color: {COLORS['divider']};
    text-align: center;
}}

QProgressBar::chunk {{
    background-color: {COLORS['primary']};
    border-radius: 4px;
}}
"""

# 欢迎语
WELCOME_TEXT = """
<html>
<head>
<style>
h1 {
    color: #2979FF;
    margin-bottom: 5px;
}
p {
    font-size: 14px;
    color: #757575;
}
</style>
</head>
<body>
<h1>欢迎使用PPT布局工具</h1>
<p>该工具可以帮您将PPT幻灯片按指定布局排版到A4纸张上。<br>
请点击左上方的「选择PPT文件」按钮开始使用。</p>
</body>
</html>
"""

# 步骤指引
STEPS_GUIDE = [
    "步骤 1: 选择PPT文件",
    "步骤 2: 设置每行PPT数量和间距",
    "步骤 3: 计算并预览布局效果",
    "步骤 4: 导出PDF文件"
] 