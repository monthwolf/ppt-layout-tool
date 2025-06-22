import os
import sys
import tempfile
import subprocess
import comtypes.client  # type: ignore
import shutil
import atexit
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

class PPTProcessor:
    """
    处理PPT文件，转换为图像并生成PDF
    """
    
    def __init__(self):
        """
        初始化PPT处理器，创建临时目录管理
        """
        # 创建一个程序级别的临时目录
        self.temp_dir = tempfile.mkdtemp(prefix="ppt_tool_")
        
        # 保存所有创建的临时文件路径
        self.temp_files = []
        
        # 注册退出时的清理函数
        atexit.register(self.cleanup_temp_files)
    
    def __del__(self):
        """
        析构函数，确保清理临时文件
        """
        self.cleanup_temp_files()
    
    def cleanup_temp_files(self):
        """
        清理所有临时文件和目录
        """
        # 清理单独创建的临时文件
        for file_path in self.temp_files:
            try:
                if os.path.exists(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"清理临时文件失败: {e}")
        
        # 清空列表
        self.temp_files = []
        
        # 清理临时目录
        try:
            if hasattr(self, 'temp_dir') and self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir, ignore_errors=True)
                print(f"已清理临时目录: {self.temp_dir}")
                self.temp_dir = None
        except Exception as e:
            print(f"清理临时目录失败: {e}")
    
    def create_temp_file(self, suffix=".png"):
        """
        创建临时文件并跟踪
        
        Args:
            suffix: 文件后缀
            
        Returns:
            临时文件路径
        """
        fd, temp_path = tempfile.mkstemp(suffix=suffix, dir=self.temp_dir)
        os.close(fd)
        self.temp_files.append(temp_path)
        return temp_path
    
    def convert_ppt_to_images(self, ppt_path, progress_callback=None):
        """
        将PPT转换为图像列表
        
        Args:
            ppt_path (str): PPT文件路径
            progress_callback (callable, optional): 进度回调函数，接收三个参数：当前进度，总进度，描述文本
            
        Returns:
            list: 图像列表
        """
        if not os.path.exists(ppt_path):
            print(f"文件不存在: {ppt_path}")
            return []
        
        # 确定PPT的格式
        is_pptx = ppt_path.lower().endswith('.pptx')
        is_ppt = ppt_path.lower().endswith('.ppt')
        
        if not (is_pptx or is_ppt):
            print(f"不支持的文件格式: {ppt_path}")
            return []
        
        # 使用PPTX库处理.pptx文件
        if is_pptx:
            try:
                return self._convert_pptx_to_images(ppt_path, progress_callback)
            except Exception as e:
                print(f"PPTX转换失败: {e}，尝试使用COM方式")
                return self._convert_ppt_via_com(ppt_path, progress_callback)
        
        # 使用COM处理.ppt文件
        if is_ppt:
            return self._convert_ppt_via_com(ppt_path, progress_callback)
        
        return []
        
    def _convert_pptx_to_images(self, pptx_path, progress_callback=None):
        """
        使用python-pptx库将PPTX转换为图像
        
        Args:
            pptx_path (str): PPTX文件路径
            progress_callback (callable, optional): 进度回调函数
        
        Returns:
            list: 图像列表
        """
        from pptx import Presentation
        
        try:
            presentation = Presentation(pptx_path)
            
            # 获取幻灯片数量
            slides = list(presentation.slides)
            slide_count = len(slides)
            
            # 如果没有幻灯片，则返回空列表
            if slide_count == 0:
                return []
            
            # 使用LibreOffice或OpenOffice转换PPTX为PDF
            pdf_path = self.create_temp_file(suffix='.pdf')
            self.temp_files.append(pdf_path)
            
            result = self._convert_to_pdf_with_libreoffice(pptx_path, pdf_path)
            if not result:
                # 如果LibreOffice转换失败，尝试使用COM接口
                return self._convert_ppt_via_com(pptx_path, progress_callback)
            
            # 将PDF转换为图片
            return self._convert_pdf_to_images(pdf_path, progress_callback)
        
        except Exception as e:
            print(f"转换PPTX时出错: {e}")
            # 如果python-pptx处理失败，尝试使用COM接口
            return self._convert_ppt_via_com(pptx_path, progress_callback)
    
    def _convert_to_pdf_with_libreoffice(self, input_path, output_path):
        """
        使用LibreOffice转换文档为PDF
        
        Args:
            input_path (str): 输入文件路径
            output_path (str): 输出PDF路径
            
        Returns:
            bool: 是否成功
        """
        try:
            # 检查是否安装了LibreOffice/OpenOffice
            if sys.platform == "win32":
                libreoffice_paths = [
                    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
                    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
                    "C:\\Program Files\\OpenOffice\\program\\soffice.exe",
                    "C:\\Program Files (x86)\\OpenOffice\\program\\soffice.exe",
                ]
                soffice_path = None
                for path in libreoffice_paths:
                    if os.path.exists(path):
                        soffice_path = path
                        break
            else:
                soffice_path = "soffice"  # 在Linux/Mac上尝试直接使用命令
            
            if not soffice_path:
                print("未找到LibreOffice/OpenOffice")
                return False
            
            # 构建命令
            cmd = [
                soffice_path,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", os.path.dirname(output_path),
                input_path
            ]
            
            # 执行命令
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
            
            # 检查是否成功
            if result.returncode != 0:
                print(f"LibreOffice转换失败: {result.stderr.decode()}")
                return False
            
            # 检查输出文件是否存在
            expected_output = os.path.join(
                os.path.dirname(output_path),
                os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
            )
            
            if os.path.exists(expected_output) and expected_output != output_path:
                # 如果输出文件名与预期不同，重命名
                shutil.move(expected_output, output_path)
            
            return os.path.exists(output_path)
        except Exception as e:
            print(f"LibreOffice转换出错: {e}")
            return False
    
    def _convert_ppt_via_com(self, ppt_path, progress_callback=None):
        """
        使用COM接口将PPT转换为图像列表
        
        Args:
            ppt_path (str): PPT文件路径
            progress_callback (callable, optional): 进度回调函数
            
        Returns:
            list: 图像列表
        """
        slide_images = []
        
        try:
            # 报告进度：准备阶段
            if progress_callback:
                progress_callback(0, 100, "正在初始化PPT转换...")
                
            # 初始化COM对象
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = True
            
            # 报告进度：正在打开PPT
            if progress_callback:
                progress_callback(10, 100, "正在打开PPT文件...")
                
            presentation = powerpoint.Presentations.Open(ppt_path)
            
            # 获取幻灯片数量
            slide_count = presentation.Slides.Count
            
            # 如果没有幻灯片，则返回空列表
            if slide_count == 0:
                powerpoint.Quit()
                return []
            
            # 报告进度：开始导出
            if progress_callback:
                progress_callback(20, 100, f"开始导出 {slide_count} 张幻灯片...")
            
            # 创建临时目录存放图像
            temp_dir = tempfile.mkdtemp(dir=self.temp_dir)
            
            # 导出PPT为图片
            for i in range(1, slide_count + 1):
                # 创建临时文件路径
                slide_path = os.path.join(temp_dir, f"slide_{i}.png")
                
                try:
                    # 导出当前幻灯片为图片
                    slide = presentation.Slides.Item(i)
                    slide.Export(slide_path, "PNG")
                    
                    # 读取导出的图像
                    image = Image.open(slide_path)
                    slide_images.append(image)
                    
                    # 报告进度
                    if progress_callback:
                        progress_callback(20 + (i * 70) // slide_count, 100, 
                                        f"已处理 {i}/{slide_count} 张幻灯片...")
                except Exception as e:
                    print(f"导出幻灯片 {i} 时出错: {e}")
                    continue
            
            # 报告进度：完成
            if progress_callback:
                progress_callback(100, 100, "PPT处理完成")
            
            # 关闭PPT
            presentation.Close()
            powerpoint.Quit()
            
            return slide_images
        
        except Exception as e:
            print(f"使用COM转换PPT时出错: {e}")
            return []
    
    def _convert_pdf_to_images(self, pdf_path, progress_callback=None):
        """
        将PDF转换为图像列表
        
        Args:
            pdf_path (str): PDF文件路径
            progress_callback (callable, optional): 进度回调函数
            
        Returns:
            list: 图像列表
        """
        from pdf2image import convert_from_path
        
        slide_images = []
        
        try:
            # 报告进度：开始准备
            if progress_callback:
                progress_callback(0, 100, "正在从PDF提取图像...")
            
            # 转换PDF为图像
            images = convert_from_path(pdf_path)
            
            # 报告进度：PDF加载完成
            if progress_callback:
                progress_callback(20, 100, f"已加载PDF，共 {len(images)} 页")
            
            # 处理每个页面
            for i, image in enumerate(images):
                # 更频繁地报告进度
                if progress_callback:
                    progress_callback(20 + (i + 1) * 70 // len(images), 100, 
                                    f"正在处理图像 {i+1}/{len(images)}...")
                
                slide_images.append(image)
            
            # 报告进度：完成
            if progress_callback:
                progress_callback(100, 100, "图像提取完成")
            
            return slide_images
        
        except Exception as e:
            print(f"转换PDF为图像时出错: {e}")
            return []
    
    def generate_pdf(self, slide_images, output_path, layout_result, config, progress_callback=None):
        """
        根据布局将PPT图像生成为PDF
        
        Args:
            slide_images: PPT幻灯片的图像列表
            output_path: 输出PDF文件路径
            layout_result: 布局计算结果
            config: 布局配置
            progress_callback (callable, optional): 进度回调函数
            
        Returns:
            布尔值，表示是否成功
        """
        temp_image_files = []  # 用于跟踪本次操作创建的临时文件
        
        try:
            # 确保有幻灯片可处理
            if not slide_images or len(slide_images) == 0:
                print("没有幻灯片可处理")
                return False
            
            # 报告进度：开始
            if progress_callback:
                progress_callback(0, 100, "准备生成PDF...")
                
            # 使用布局计算结果中的页面尺寸（包含方向信息）
            page_width_mm = layout_result["page_width"]
            page_height_mm = layout_result["page_height"]
            
            # 创建PDF页面尺寸
            page_size = (page_width_mm * mm, page_height_mm * mm)
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 报告进度：准备字体
            if progress_callback:
                progress_callback(10, 100, "正在准备字体...")
            
            # 注册中文字体用于页码显示
            chinese_font_name = "Helvetica"  # 默认字体
            try:
                # 首先尝试使用项目自带的Source Han Sans字体
                source_han_sans_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 
                                                  'resources', 'SourceHanSans.ttf')
                
                if os.path.exists(source_han_sans_path):
                    chinese_font_name = "SourceHanSans"
                    pdfmetrics.registerFont(TTFont(chinese_font_name, source_han_sans_path))
                    print(f"已注册内置中文字体用于页码显示: {chinese_font_name}")
                else:
                    # 尝试使用系统中文字体
                    font_paths = [
                        "C:/Windows/Fonts/simhei.ttf",
                        "C:/Windows/Fonts/msyh.ttc",
                        "/System/Library/Fonts/PingFang.ttc"
                    ]
                    for path in font_paths:
                        if os.path.exists(path):
                            font_name = "SimHei" if "simhei" in path.lower() else "ChineseFont"
                            pdfmetrics.registerFont(TTFont(font_name, path))
                            chinese_font_name = font_name
                            print(f"已注册系统字体用于页码显示: {chinese_font_name}")
                            break
            except Exception as e:
                print(f"注册中文字体失败: {e}")
            
            # 创建PDF画布
            c = canvas.Canvas(output_path, pagesize=page_size)
            
            # 获取布局参数
            rows = layout_result["rows"]
            columns = layout_result["columns"]
            item_width_mm = layout_result["item_width"]
            item_height_mm = layout_result["item_height"]
            h_spacing_mm = config["h_spacing"]
            v_spacing_mm = config["v_spacing"]
            margin_left_mm = config["margin_left"]
            margin_top_mm = config["margin_top"]
            margin_bottom_mm = config.get("margin_bottom", margin_top_mm)
            margin_right_mm = config.get("margin_right", margin_left_mm)
            
            # 获取页码显示设置
            show_ppt_numbers = config.get("show_ppt_numbers", True)
            show_page_numbers = config.get("show_page_numbers", True)
            
            # 计算每页可以放置的幻灯片数量
            items_per_page = rows * columns
            
            # 处理每一页
            page_count = 0
            slide_count = len(slide_images)
            total_pages = (slide_count + items_per_page - 1) // items_per_page
            
            # 报告进度：开始生成页面
            if progress_callback:
                progress_callback(20, 100, f"开始生成 {total_pages} 页PDF...")

            for page_idx in range(total_pages):
                if page_idx > 0:
                    # 创建新页面
                    c.showPage()
                
                page_count += 1
                
                # 报告当前页面进度
                if progress_callback:
                    progress_callback(20 + (page_idx * 70) // total_pages, 100,
                                    f"正在生成第 {page_idx + 1}/{total_pages} 页...")

                # 处理当前页的每个位置
                for pos in range(items_per_page):
                    slide_idx = page_idx * items_per_page + pos
                    
                    if slide_idx >= slide_count:
                        break
                    
                    # 计算行和列
                    row = pos // columns
                    col = pos % columns
                    
                    # 计算位置（毫米）
                    x_mm = margin_left_mm + col * (item_width_mm + h_spacing_mm)
                    y_mm = page_height_mm - margin_top_mm - item_height_mm - row * (item_height_mm + v_spacing_mm)
                    
                    # 转换为点（PDF单位）
                    x = x_mm * mm
                    y = y_mm * mm
                    width = item_width_mm * mm
                    height = item_height_mm * mm
                    
                    try:
                        # 获取当前幻灯片图像
                        slide_img = slide_images[slide_idx]
                        
                        # 创建临时文件保存调整大小后的图像
                        tmp_path = self.create_temp_file(suffix=".png")
                        temp_image_files.append(tmp_path)
                        
                        # 调整图像大小
                        slide_img.save(tmp_path, format="PNG")
                        
                        # 将图像添加到PDF
                        c.drawImage(tmp_path, x, y, width, height)
                        
                        # 添加PPT定位页码标记
                        if show_ppt_numbers:
                            c.setFont(chinese_font_name, 8)
                            label = f"{page_idx+1}-{pos+1}"
                            # 放在PPT的左下角
                            c.drawString(x, y - 10, label)
                    except Exception as e:
                        print(f"处理幻灯片 {slide_idx+1} 时出错: {e}")
                        # 如果单个幻灯片处理失败，继续处理下一个
                        continue
                
                # 添加纸张页码（在页面右下角）
                if show_page_numbers:
                    c.setFont(chinese_font_name, 10)
                    # 使用中文页码格式
                    page_number_text = f"第 {page_idx+1} 页"
                    # 计算页码位置在右下角
                    page_number_x = page_width_mm * mm - margin_right_mm * mm - 25  # 增加空间以容纳中文
                    page_number_y = margin_bottom_mm * mm 
                    c.drawString(page_number_x, page_number_y, page_number_text)
            
            # 报告进度：正在保存
            if progress_callback:
                progress_callback(95, 100, "正在保存PDF文件...")
            
            # 保存PDF
            c.save()
            
            # 报告进度：完成
            if progress_callback:
                progress_callback(100, 100, "PDF生成完成")
            
            return True
        
        except Exception as e:
            print(f"生成PDF时发生错误: {e}")
            if progress_callback:
                progress_callback(100, 100, f"错误: {e}")
            return False
        finally:
            # 清理本次操作的临时文件
            for temp_file in temp_image_files:
                try:
                    if os.path.exists(temp_file):
                        os.unlink(temp_file)
                except Exception as e:
                    print(f"清理临时文件失败: {temp_file} - {e}")
    
    def generate_pdf_with_index(self, markdown_text, content_pdf_path, final_output_path, progress_callback=None):
        """
        将Markdown索引和内容PDF合并
        
        Args:
            markdown_text (str): Markdown格式的索引
            content_pdf_path (str): 内容PDF的路径
            final_output_path (str): 最终输出路径
            progress_callback (callable, optional): 进度回调函数
            
        Returns:
            bool: 是否成功
        """
        # 创建一个临时的PDF文件用于存放索引
        index_fd, index_pdf_path = tempfile.mkstemp(suffix=".pdf")
        os.close(index_fd)

        try:
            # 报告进度：开始准备
            if progress_callback:
                progress_callback(0, 100, "正在准备生成索引PDF...")
            
            # 获取内容PDF的页面方向和尺寸
            content_pdf_config = self._get_pdf_config(content_pdf_path)
            if not content_pdf_config:
                print("警告：无法获取内容PDF的配置信息，将使用默认A4尺寸")
                content_pdf_config = {"pagesize": A4}
            
            # 报告进度：Markdown处理
            if progress_callback:
                progress_callback(20, 100, "正在将Markdown转换为PDF...")
                
            # 1. 将Markdown转换为PDF，使用相同的页面尺寸
            result = self._markdown_to_pdf(markdown_text, index_pdf_path, content_pdf_config)
            if not result:
                print("转换Markdown到PDF失败")
                return False

            # 报告进度：合并PDF
            if progress_callback:
                progress_callback(60, 100, "正在合并索引与内容PDF...")
                
            # 2. 合并PDF
            try:
                merger = PdfWriter()
                
                # 首先添加索引PDF
                with open(index_pdf_path, "rb") as f:
                    index_pdf = PdfReader(f)
                    merger.append(index_pdf)

                # 然后添加内容PDF
                with open(content_pdf_path, "rb") as f:
                    content_pdf = PdfReader(f)
                    merger.append(content_pdf)

                # 报告进度：写入最终PDF
                if progress_callback:
                    progress_callback(80, 100, "正在写入最终PDF文件...")
                    
                # 写入最终文件
                with open(final_output_path, "wb") as f:
                    merger.write(f)
                
                # 报告进度：完成
                if progress_callback:
                    progress_callback(100, 100, "PDF生成完成")
                
                merger.close()
                return True
            except Exception as e:
                print(f"合并PDF时出错: {e}")
                return False

        except Exception as e:
            print(f"生成索引PDF时出错: {e}")
            return False
        finally:
            # 清理临时的索引PDF
            if os.path.exists(index_pdf_path):
                os.unlink(index_pdf_path)
    
    def _get_pdf_config(self, pdf_path):
        """
        获取PDF的配置信息，包括页面尺寸和方向
        
        Args:
            pdf_path (str): PDF文件路径
            
        Returns:
            dict: 包含pagesize等信息的配置字典
        """
        try:
            with open(pdf_path, "rb") as f:
                pdf = PdfReader(f)
                if len(pdf.pages) > 0:
                    page = pdf.pages[0]
                    # 获取页面尺寸（以点为单位）
                    width = float(page.mediabox.width)
                    height = float(page.mediabox.height)
                    
                    # 转换为mm
                    width_mm = width * 0.352778
                    height_mm = height * 0.352778
                    
                    # 判断页面方向
                    is_landscape = width > height
                    
                    # 创建页面尺寸元组
                    if is_landscape:
                        pagesize = landscape(A4) if abs(width_mm - 297) < 5 and abs(height_mm - 210) < 5 else (width, height)
                    else:
                        pagesize = A4 if abs(width_mm - 210) < 5 and abs(height_mm - 297) < 5 else (width, height)
                    
                    return {
                        "pagesize": pagesize,
                        "is_landscape": is_landscape,
                        "width": width,
                        "height": height
                    }
            
            return None
        except Exception as e:
            print(f"获取PDF配置时出错: {e}")
            return None

    def _markdown_to_pdf(self, markdown_text, output_path, content_pdf_config=None):
        """
        将Markdown文本转换为PDF文件
        
        Args:
            markdown_text (str): Markdown文本
            output_path (str): 输出PDF路径
            content_pdf_config (dict, optional): 内容PDF的配置，包括pagesize等信息
        """
        # 注册中文字体
        try:
            # 首先尝试使用项目自带的Source Han Sans字体
            source_han_sans_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 
                                               'resources', 'SourceHanSans.ttf')
            
            if os.path.exists(source_han_sans_path):
                chinese_font_name = "SourceHanSans"
                pdfmetrics.registerFont(TTFont(chinese_font_name, source_han_sans_path))
                print(f"已注册内置中文字体: {chinese_font_name}")
            else:
                # 如果内置字体不可用，尝试使用系统字体
                font_paths = [
                    # Windows 中文字体路径
                    "C:/Windows/Fonts/simhei.ttf",           # 黑体
                    "C:/Windows/Fonts/simsun.ttc",           # 宋体
                    "C:/Windows/Fonts/simkai.ttf",           # 楷体
                    "C:/Windows/Fonts/msyh.ttc",             # 微软雅黑
                    # Mac OS 中文字体路径
                    "/System/Library/Fonts/PingFang.ttc",    # 苹方
                    # Linux 中文字体路径
                    "/usr/share/fonts/truetype/arphic/uming.ttc",
                    "/usr/share/fonts/wqy-microhei/wqy-microhei.ttc"
                ]
                
                # 搜索系统中的字体文件
                font_file = None
                for path in font_paths:
                    if os.path.exists(path):
                        font_file = path
                        break
                
                # 如果找到字体文件，则注册字体
                if font_file:
                    chinese_font_name = "SimHei" if "simhei" in font_file.lower() else "ChineseFont"
                    pdfmetrics.registerFont(TTFont(chinese_font_name, font_file))
                    print(f"已注册系统中文字体: {chinese_font_name} 从 {font_file}")
                else:
                    print("警告: 未找到合适的中文字体文件")
                    chinese_font_name = "Helvetica"  # 回退到默认字体
        except Exception as e:
            print(f"注册中文字体时出错: {e}")
            chinese_font_name = "Helvetica"

        # 使用内容PDF的页面尺寸
        pagesize = None
        if content_pdf_config and "pagesize" in content_pdf_config:
            pagesize = content_pdf_config["pagesize"]
        else:
            pagesize = A4  # 默认使用A4尺寸
        
        # 设置文档和样式
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        
        doc = SimpleDocTemplate(output_path, pagesize=pagesize)
        styles = getSampleStyleSheet()
        
        # 修改现有样式以支持中文，而不是添加新样式
        styles['Title'].fontName = chinese_font_name
        styles['Heading1'].fontName = chinese_font_name
        styles['Heading2'].fontName = chinese_font_name
        styles['Heading3'].fontName = chinese_font_name
        styles['BodyText'].fontName = chinese_font_name
        styles['Normal'].fontName = chinese_font_name
        
        # 创建列表项样式（这个样式在默认样式表中可能不存在）
        bullet_style = ParagraphStyle(
            'ChineseBullet',
            parent=styles['Normal'],
            leftIndent=20,
            fontName=chinese_font_name
        )
        
        # 使用更简单的方法解析Markdown
        # 先将markdown文本分割成段落
        paragraphs = markdown_text.split('\n\n')
        story = []
        
        for paragraph in paragraphs:
            paragraph = paragraph.strip()
            if not paragraph:
                continue
                
            # 处理标题
            if paragraph.startswith('#'):
                heading_level = 0
                while paragraph.startswith('#'):
                    heading_level += 1
                    paragraph = paragraph[1:]
                paragraph = paragraph.strip()
                
                if heading_level == 1:
                    story.append(Paragraph(paragraph, styles['Title']))
                elif heading_level == 2:
                    story.append(Paragraph(paragraph, styles['Heading1']))
                elif heading_level == 3:
                    story.append(Paragraph(paragraph, styles['Heading2']))
                else:
                    story.append(Paragraph(paragraph, styles['Heading3']))
            
            # 处理列表
            elif paragraph.startswith('* ') or paragraph.startswith('- '):
                lines = paragraph.split('\n')
                for line in lines:
                    line = line.strip()
                    if line.startswith('* ') or line.startswith('- '):
                        line = line[2:].strip()
                        story.append(Paragraph(line, bullet_style, bulletText='•'))
            
            # 普通段落
            else:
                story.append(Paragraph(paragraph, styles['BodyText']))
                
            # 添加间距
            story.append(Spacer(1, 6))
        
        # 创建PDF
        doc.build(story)
        print(f"Markdown已成功转换为PDF: {output_path}")
        return True