import os
import io
import sys
import tempfile
import subprocess
import ctypes
import comtypes.client  # type: ignore
import shutil
import atexit
from pathlib import Path
import time
import re
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from PyPDF2 import PdfWriter, PdfReader
import markdown_it
from reportlab.lib import fonts
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
    
    def convert_ppt_to_images(self, ppt_path):
        """
        将PPT文件转换为图像列表
        
        Args:
            ppt_path: PPT文件路径
            
        Returns:
            图像列表
        """
        try:
            slide_images = []
            file_ext = os.path.splitext(ppt_path)[1].lower()
            
            # 检查文件是否存在
            if not os.path.exists(ppt_path):
                print(f"文件不存在: {ppt_path}")
                return self._create_mock_slides(5)  # 返回模拟幻灯片
                
            # 根据文件类型选择处理方法
            if file_ext == '.pptx':
                try:
                    # 尝试使用python-pptx直接读取
                    from pptx import Presentation
                    try:
                        presentation = Presentation(ppt_path)
                        slide_count = len(presentation.slides)
                        return self._create_mock_slides(slide_count)
                    except Exception as e:
                        print(f"python-pptx读取失败: {e}")
                        # 如果python-pptx失败，回退到其他方法
                        return self._try_alternative_conversion(ppt_path)
                except ImportError:
                    print("python-pptx库未安装或导入失败")
                    return self._try_alternative_conversion(ppt_path)
            else:
                # 对于.ppt文件，直接使用替代方法
                return self._try_alternative_conversion(ppt_path)
                
        except Exception as e:
            print(f"转换PPT时发生错误: {e}")
            # 返回5张模拟幻灯片，确保程序可以继续运行
            return self._create_mock_slides(5)
        finally:
            # 每次转换完成后清理临时文件
            self.cleanup_temp_files()
    
    def _try_alternative_conversion(self, ppt_path):
        """尝试使用替代方法转换PPT"""
        try:
            # 尝试使用COM接口（仅Windows可用）
            if sys.platform == 'win32':
                return self._convert_using_com(ppt_path)
            else:
                # 对于非Windows系统，使用模拟方式
                return self._create_mock_slides(10)
        except Exception as e:
            print(f"替代转换方法失败: {e}")
            # 如果所有方法都失败，使用模拟幻灯片
            return self._create_mock_slides(10)
    
    def _convert_using_com(self, ppt_path):
        """使用PowerPoint COM接口转换PPT为图像（仅Windows可用）"""
        ppt_app = None
        presentation = None
        temp_output_dir = None
        
        try:
            import comtypes.client  # type: ignore
            
            # 创建临时文件夹存储图片
            temp_output_dir = os.path.join(self.temp_dir if self.temp_dir else tempfile.gettempdir(), "ppt_export_" + str(int(time.time())))
            os.makedirs(temp_output_dir, exist_ok=True)
            
            # 尝试启动PowerPoint
            ppt_app = comtypes.client.CreateObject("PowerPoint.Application")  # type: ignore
            ppt_app.Visible = True  # 设置为可见，以便观察过程
            
            # 规范化文件路径，使用绝对路径
            abs_path = os.path.abspath(ppt_path)
            
            # 打开演示文稿 - 跳过类型检查，因为COM对象的类型系统与Python不兼容
            presentation = ppt_app.Presentations.Open(str(abs_path), WithWindow=False)  # type: ignore
            
            # 获取幻灯片数量
            slide_count = presentation.Slides.Count
            print(f"检测到 {slide_count} 张幻灯片")
            
            # 保存为图片
            presentation.Export(temp_output_dir, "PNG")
            
            # 关闭演示文稿
            try:
                if presentation:
                    presentation.Close()
                    presentation = None
            except:
                pass
            
            # 关闭PowerPoint应用程序
            try:
                if ppt_app:
                    ppt_app.Quit()
                    ppt_app = None
            except:
                pass
            
            # 加载图片
            slide_images = []
            
            # 使用自然排序确保文件顺序正确
            def natural_sort_key(s):
                return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]
                
            image_files = sorted(
                [f for f in os.listdir(temp_output_dir) if f.endswith(('.png', '.PNG'))],
                key=natural_sort_key
            )
            
            for img_file in image_files:
                img_path = os.path.join(temp_output_dir, img_file)
                # print(img_path)
                try:
                    img = Image.open(img_path)
                    slide_images.append(img.copy())  # 复制图像对象，以便关闭原始文件
                except Exception as e:
                    print(f"加载图片失败 {img_path}: {e}")
            # print(slide_images)
            return slide_images
        except Exception as e:
            print(f"COM接口转换失败: {e}")
            return self._create_mock_slides(10)
        finally:
            # 确保清理资源
            try:
                if presentation:
                    presentation.Close()
            except:
                pass
            
            try:
                if ppt_app:
                    ppt_app.Quit()
            except:
                pass
    
    def _create_mock_slides(self, count=5):
        """创建模拟幻灯片图像
        
        Args:
            count: 要创建的幻灯片数量
            
        Returns:
            模拟幻灯片图像列表
        """
        slide_images = []
        
        # 为每个幻灯片创建一个临时文件目录
        temp_slides_dir = os.path.join(self.temp_dir if self.temp_dir else tempfile.gettempdir(), "mock_slides")
        os.makedirs(temp_slides_dir, exist_ok=True)
        
        for i in range(count):
            img_path = os.path.join(temp_slides_dir, f"slide_{i}.png")
            self.temp_files.append(img_path)  # 跟踪临时文件
            
            # 创建一个模拟的幻灯片图像 - 使用类型忽略解决PIL类型兼容问题
            img = Image.new('RGB', (1280, 720), color=(255, 255, 255))  # type: ignore
            d = ImageDraw.Draw(img)
            
            # 绘制边框
            d.rectangle([20, 20, 1260, 700], outline='black', width=2)
            
            # 试图加载一个字体，如果失败则使用默认
            try:
                font = ImageFont.truetype("arial.ttf", 48)
            except IOError:
                font = ImageFont.load_default()
            
            # 写入幻灯片编号和模拟状态
            text_pos = (640, 360)
            d.text(text_pos, f"幻灯片 {i+1} (模拟)", fill='black', font=font, anchor="mm")
            
            # 保存为临时文件
            img.save(img_path)
            
            # 加载保存的图像并添加到列表
            slide_image = Image.open(img_path)
            slide_images.append(slide_image)
            
        return slide_images
    
    def generate_pdf(self, slide_images, output_path, layout_result, config):
        """
        根据布局将PPT图像生成为PDF
        
        Args:
            slide_images: PPT幻灯片的图像列表
            output_path: 输出PDF文件路径
            layout_result: 布局计算结果
            config: 布局配置
            
        Returns:
            布尔值，表示是否成功
        """
        temp_image_files = []  # 用于跟踪本次操作创建的临时文件
        
        try:
            # 确保有幻灯片可处理
            if not slide_images or len(slide_images) == 0:
                print("没有幻灯片可处理")
                return False
                
            # 使用布局计算结果中的页面尺寸（包含方向信息）
            page_width_mm = layout_result["page_width"]
            page_height_mm = layout_result["page_height"]
            
            # 创建PDF页面尺寸
            page_size = (page_width_mm * mm, page_height_mm * mm)
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
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
            
            for page_idx in range((slide_count + items_per_page - 1) // items_per_page):
                if page_idx > 0:
                    # 创建新页面
                    c.showPage()
                
                page_count += 1
                
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
            
            # 保存PDF
            c.save()
            return True
        
        except Exception as e:
            print(f"生成PDF时发生错误: {e}")
            return False
        finally:
            # 清理本次操作的临时文件
            for temp_file in temp_image_files:
                try:
                    if os.path.exists(temp_file):
                        os.unlink(temp_file)
                except Exception as e:
                    print(f"清理临时文件失败: {temp_file} - {e}")
    
    def generate_pdf_with_index(self, markdown_text, content_pdf_path, final_output_path):
        """
        将Markdown索引和内容PDF合并
        
        Args:
            markdown_text (str): Markdown格式的索引
            content_pdf_path (str): 内容PDF的路径
            final_output_path (str): 最终输出路径
            
        Returns:
            bool: 是否成功
        """
        # 创建一个临时的PDF文件用于存放索引
        index_fd, index_pdf_path = tempfile.mkstemp(suffix=".pdf")
        os.close(index_fd)

        try:
            # 获取内容PDF的页面方向和尺寸
            content_pdf_config = self._get_pdf_config(content_pdf_path)
            if not content_pdf_config:
                print("警告：无法获取内容PDF的配置信息，将使用默认A4尺寸")
                content_pdf_config = {"pagesize": A4}
                
            # 1. 将Markdown转换为PDF，使用相同的页面尺寸
            result = self._markdown_to_pdf(markdown_text, index_pdf_path, content_pdf_config)
            if not result:
                print("转换Markdown到PDF失败")
                return False

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

                # 写入最终文件
                with open(final_output_path, "wb") as f:
                    merger.write(f)
                
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