class LayoutCalculator:
    """
    计算PPT在A4页面上的布局
    """
    
    def calculate_layout(self, slide_images, config):
        """
        计算PPT在A4页面上的最佳布局
        
        Args:
            slide_images: PPT幻灯片的图像列表
            config: 布局配置字典
            
        Returns:
            字典，包含布局信息
        """
        if not slide_images:
            return {
                "rows": 0,
                "columns": config["columns"],
                "item_width": 0,
                "item_height": 0,
                "pages_needed": 0,
                "is_landscape": config.get("is_landscape", True)
            }
        
        # 获取第一张幻灯片的宽高比作为参考
        # 获取PPT图像的原始尺寸
        sample_image = slide_images[0]
        orig_width, orig_height = sample_image.size
        aspect_ratio = orig_width / orig_height
        
        # 获取A4页面尺寸设置
        std_page_width = 210  # A4宽度(mm)
        std_page_height = 297  # A4高度(mm)
        
        # 根据用户选择设置页面方向
        is_landscape = config.get("is_landscape", True)  # 默认横向
        
        if is_landscape:
            # 横向A4 (297×210mm)
            page_width = std_page_height   # 宽度为297mm
            page_height = std_page_width   # 高度为210mm
        else:
            # 纵向A4 (210×297mm)
            page_width = std_page_width    # 宽度为210mm
            page_height = std_page_height  # 高度为297mm
        
        # 计算可用区域，考虑边距
        available_width = page_width - config["margin_left"] - config["margin_right"]
        available_height = page_height - config["margin_top"] - config["margin_bottom"]
        
        # 根据列数计算每个PPT的宽度
        columns = config["columns"]
        h_spacing = config["h_spacing"]
        total_h_spacing = h_spacing * (columns - 1)
        
        # 计算PPT在页面上的最大宽度
        item_width = (available_width - total_h_spacing) / columns
        item_height = item_width / aspect_ratio
        
        # 计算能放几行
        v_spacing = config["v_spacing"]
        rows = int((available_height + v_spacing) / (item_height + v_spacing))
        
        # 确保至少有一行
        rows = max(1, rows)
        
        # 计算需要多少页
        items_per_page = rows * columns
        total_slides = len(slide_images)
        pages_needed = (total_slides + items_per_page - 1) // items_per_page
        
        return {
            "rows": rows,
            "columns": columns,
            "item_width": item_width,
            "item_height": item_height,
            "aspect_ratio": aspect_ratio,
            "pages_needed": pages_needed,
            "items_per_page": items_per_page,
            "is_landscape": is_landscape,
            "page_width": page_width,
            "page_height": page_height
        } 