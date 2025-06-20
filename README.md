# PPT布局工具

一个用于处理PPT文件并将其按照特定布局排版到A4纸张上的工具，采用现代UI设计。

## 功能

- 导入PPT文件，将每张幻灯片转换为图片
- 支持新版PPTX和旧版PPT文件格式
- 根据用户设置的每行PPT数量，自动计算最佳布局
- 在界面上显示布局预览效果
- 支持调整PPT间水平和垂直间距
- 支持自定义页边距设置
- 支持横向和纵向A4页面排列选择
- 按步骤引导用户完成全部操作
- 导出为PDF文件，每张PPT带有页码标记
- 自动清理所有临时文件，确保不占用额外磁盘空间

## 系统要求

- Windows系统（建议Windows 10及以上）
- 安装了Anaconda或Miniconda环境
- 程序会自动创建所需的Python环境
- 处理旧版PPT(.ppt)文件需要安装Microsoft PowerPoint

## 快速开始

1. 确保已安装Anaconda或Miniconda
2. 双击`start.bat`启动程序，它会自动：
   - 创建名为`ppt_tool`的Conda环境
   - 安装所需的Python依赖库
   - 启动PPT布局工具
   - 如果主程序启动失败，会自动尝试启动备用版本
3. 按照界面上的步骤引导操作：
   - 步骤1：选择PPT文件
   - 步骤2：设置布局参数
   - 步骤3：查看预览效果
   - 步骤4：导出PDF文件

## 支持的PPT格式

### PPTX格式（推荐）
- 使用python-pptx库直接读取
- 不需要额外软件支持

### PPT格式（旧版）
- 使用Windows PowerPoint COM接口处理
- **需要**安装Microsoft PowerPoint
- 如未安装PowerPoint，将使用模拟幻灯片代替

## 常见问题与解决方案

### PPT文件无法正确读取

如果出现"Package not found"或其他PPT读取错误：

1. **对于PPTX文件**:
   - 确保文件未损坏，可以尝试在PowerPoint中打开并重新保存

2. **对于旧版PPT文件**:
   - 确保已安装Microsoft PowerPoint
   - 尝试将PPT文件保存为PPTX格式
   - 避免使用包含特殊字符的路径

3. **所有文件**:
   - 避免使用过长或包含特殊字符的文件路径
   - 将文件移动到较短的英文路径下

### DLL加载失败问题

如果遇到类似 `DLL load failed while importing QtCore` 的错误，请尝试以下解决方案：

1. **安装Visual C++ Redistributable**
   - 下载并安装最新的[Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)

2. **使用备用启动脚本**
   - 双击`start_alt.bat`，它会使用PySide6代替PyQt6启动程序
   - 此版本UI略有不同，但功能相同

3. **手动安装依赖**
   - 如果问题仍未解决，请尝试在命令行中运行以下命令：
   ```
   pip uninstall -y PyQt6 PyQt6-Qt6 PyQt6-sip pyqt6-tools
   pip install PyQt6==6.4.2 PyQt6-Qt6==6.4.2 PyQt6-sip==13.5.1
   ```

### 包版本冲突

如果安装依赖时出现版本冲突错误：

1. **检查是否有多余的旧版本包**
   ```
   pip list | findstr PyQt
   ```

2. **彻底移除所有相关包后重新安装**
   ```
   pip uninstall -y PyQt6 PyQt6-Qt6 PyQt6-sip pyqt6-tools
   pip install -r requirements.txt
   ```

3. **如果仍有问题，尝试直接使用备用版本**
   ```
   start_alt.bat
   ```

### 环境相关问题

如果环境创建过程中遇到问题：

1. **更新Conda**
   ```
   conda update -n base -c defaults conda
   ```

2. **手动创建环境**
   ```
   conda create -n ppt_tool python=3.9
   conda activate ppt_tool
   pip install -r requirements.txt
   ```

## 注意事项

- 使用横向A4纸张作为默认布局
- 每个PPT会按原始宽高比进行缩放
- 每页PPT标记格式为"页码-位置号"
- 首次运行可能需要较长时间创建环境和安装依赖
- 处理旧版PPT(.ppt)格式需要安装Microsoft PowerPoint
- 本程序使用PyQt6 6.4.2版本，与pyqt6-tools 6.4.2.3.3保持版本兼容

## 技术细节

本工具使用以下技术：
- PyQt6/PySide6 - 现代化UI界面
- python-pptx - 读取PPTX文件
- comtypes - 通过COM接口处理旧版PPT文件
- Pillow - 图像处理
- reportlab - PDF生成

## 资源管理

### 临时文件处理
- 程序使用系统临时目录存储PPT转换的临时图像
- 在以下情况下会自动清理临时文件：
  - 每次成功导出PDF后
  - 程序正常关闭时
  - 发生错误或异常时
  - 转换新PPT文件前
- 即使程序意外崩溃，临时文件也通常会在系统重启后被清理 