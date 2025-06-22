"""
定义应用程序的现代UI样式
"""

# 应用程序的主题颜色
COLORS = {
    "primary": "#4A90E2",         # Softer Blue
    "primary_dark": "#357ABD",    # Darker shade for hover
    "success": "#50E3C2",         # Mint Green
    "error": "#E94E77",           # Soft Red
    "background": "#FDFDFD",      # Off-white background
    "surface": "#FFFFFF",         # Pure white for surfaces
    "text_primary": "#333333",    # Dark Gray
    "text_secondary": "#777777",  # Medium Gray
    "divider": "#EAEAEA",         # Light Gray
}

# 应用程序样式表
STYLESHEET = f"""
    QMainWindow, QWidget {{
        background-color: {COLORS['background']};
        color: {COLORS['text_primary']};
        font-family: "Segoe UI", "Helvetica Neue", "Arial", sans-serif;
    }}
    
    QGroupBox {{
        background-color: {COLORS['surface']};
        border: 1px solid {COLORS['divider']};
        overflow-y: scroll;
        border-radius: 8px;
        font-size: 14px;
        font-weight: bold;
        margin-top: 10px;
        padding: 20px;
    }}
    
    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 0 10px;
        left: 10px;
        color: {COLORS['primary']};
    }}

    QPushButton {{
        background-color: {COLORS['surface']};
        border: 1px solid {COLORS['divider']};
        border-radius: 5px;
        padding: 10px 15px;
        font-size: 13px;
        color: {COLORS['text_primary']};
        transition: all 0.3s ease;
    }}
    
    QPushButton:hover {{
        background-color: #F0F0F0;
        border-color: #C0C0C0;
    }}
    
    QPushButton:pressed {{
        background-color: {COLORS['divider']};
    }}
    
    QPushButton:disabled {{
        background-color: #F8F8F8;
        color: #BBBBBB;
        border-color: #E0E0E0;
    }}
    
    QPushButton#accentButton {{
        background-color: {COLORS['primary']};
        color: white;
        font-weight: bold;
        border: none;
    }}
    
    QPushButton#accentButton:hover {{
        background-color: {COLORS['primary_dark']};
    }}
    
    QPushButton#accentButton:disabled {{
        background-color: #A0C0E0;
        color: white;
        opacity: 0.7;
    }}

    QLabel {{
        font-size: 13px;
        color: {COLORS['text_secondary']};
    }}

    QLabel#infoLabel {{
        color: {COLORS['text_secondary']};
        font-style: italic;
    }}
    
    QSpinBox, QDoubleSpinBox, QTextEdit {{
        background-color: {COLORS['surface']};
        border: 1px solid {COLORS['divider']};
        border-radius: 5px;
        padding: 8px;
        font-size: 13px;
    }}

    QSpinBox:focus, QDoubleSpinBox:focus, QTextEdit:focus {{
        border: 1px solid {COLORS['primary']};
    }}

    QRadioButton::indicator {{
        width: 16px;
        height: 16px;
    }}

    QCheckBox::indicator {{
        width: 16px;
        height: 16px;
    }}
    
    QStatusBar {{
        font-size: 12px;
    }}
    
    QScrollBar:vertical {{
        border: none;
        background: {COLORS['background']};
        width: 10px;
        margin: 0px 0px 0px 0px;
    }}
    QScrollBar::handle:vertical {{
        background: {COLORS['divider']};
        min-height: 20px;
        border-radius: 5px;
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
    }}
    QScrollBar:horizontal {{
        border: none;
        background: {COLORS['background']};
        height: 10px;
        margin: 0px 0px 0px 0px;
    }}
    QScrollBar::handle:horizontal {{
        background: {COLORS['divider']};
        min-width: 20px;
        border-radius: 5px;
    }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0px;
    }}
"""

# 欢迎语
WELCOME_TEXT = """
<div style="text-align: center; padding: 20px;">
  <h1 style="color: #333; font-size: 28px; font-weight: 600; margin-bottom: 10px;">
    欢迎使用 PPT 布局工具
  </h1>
  <p style="color: #555; font-size: 15px; max-width: 450px; margin: 0 auto 25px auto;">
    一个现代、高效的工具，旨在将您的演示文稿轻松排版为专业、整洁的PDF文档。
  </p>
  <div style="text-align: center; max-width: 400px; margin: 0 auto; line-height: 1.8;">
    <p><b>✓ 灵活布局：</b>自由定义每页的行列数量与间距。</p>
    <p><b>✓ 智能预览：</b>在导出前实时查看最终布局效果。</p>
    <p><b>✓ AI 索引：</b>一键生成提示词，利用AI创建专业索引。</p>
  </div>
  <p style="color: #888; font-size: 13px; margin-top: 30px;">
    请点击左上角的"选择PPT文件"按钮，开始您的创作之旅。
  </p>
</div>
"""

# 步骤指引
STEPS_GUIDE = [
    "选择文件",
    "布局设置",
    "效果预览",
    "导出文件",
    "AI索引"
] 