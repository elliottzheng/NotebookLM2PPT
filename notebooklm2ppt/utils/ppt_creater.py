"""从 PaddleOCR (PP-Structure) JSON 直接创建 PPT"""

import os
import json
import argparse
import copy
from pathlib import Path
import numpy as np
from PIL import Image
from notebooklm2ppt.pdf2png import pdf_to_png
from notebooklm2ppt.utils.ppt_combiner import clean_ppt
from notebooklm2ppt.utils.edge_diversity import compute_edge_diversity_numpy
from spire.presentation import *
from spire.presentation.common import *

# ============================================================================
# 文本分析工具函数
# ============================================================================


def calculate_font_size(height,
                        min_font_size=8,
                        is_multiline=False,
                        line_count=1):
    """
    根据文本框大小和文本内容计算合适的字体大小
    
    Args:
        text: 文本内容
        width: 文本框宽度
        height: 文本框高度
        min_font_size: 最小字体大小
        is_multiline: 是否多行文本
        line_count: 行数
        
    Returns:
        int: 计算后的字体大小
    """
    if is_multiline and line_count > 1:
        line_height = height / line_count
        font_size = line_height * 0.75
    else:
        font_size = height * 0.75

    font_size = max(min_font_size, int(font_size))
    return font_size


def get_line_count(block_bbox, ocr_boxes):
    """
    计算文本块内的实际行数
    
    Args:
        block_bbox: 文本块边界框 [x1, y1, x2, y2]
        ocr_boxes: OCR识别的所有文本框列表
        
    Returns:
        int: 文本块内的行数
    """
    bx1, by1, bx2, by2 = block_bbox

    # 筛选出该 block 范围内的 OCR 框 (使用中心点判定)
    contained_boxes = []
    for obox in ocr_boxes:
        ox1, oy1, ox2, oy2 = obox
        cx = (ox1 + ox2) / 2
        cy = (oy1 + oy2) / 2

        if bx1 <= cx <= bx2 and by1 <= cy <= by2:
            contained_boxes.append(obox)

    if not contained_boxes:
        return 0

    # 按 y 中心点排序并进行简单的聚类分析行数
    contained_boxes.sort(key=lambda b: (b[1] + b[3]) / 2)

    line_count = 1
    last_y_center = (contained_boxes[0][1] + contained_boxes[0][3]) / 2
    last_h = contained_boxes[0][3] - contained_boxes[0][1]

    for j in range(1, len(contained_boxes)):
        curr_y_center = (contained_boxes[j][1] + contained_boxes[j][3]) / 2
        curr_h = contained_boxes[j][3] - contained_boxes[j][1]

        # 阈值：如果垂直间距超过行高的 60%，判定为新行
        if abs(curr_y_center - last_y_center) > max(last_h, curr_h) * 0.6:
            line_count += 1
            last_y_center = curr_y_center
            last_h = curr_h

    return line_count


# ============================================================================
# PPT设置函数
# ============================================================================


def setup_presentation(pdf_size):
    """
    根据PDF尺寸创建并设置PPT
    
    Args:
        pdf_width: PDF宽度
        pdf_height: PDF高度
        
    Returns:
        tuple: (presentation对象, ppt_width, ppt_height, scale缩放比例)
    """
    presentation = Presentation()
    if presentation.Slides.Count > 0:
        presentation.Slides.RemoveAt(0)

    pdf_width, pdf_height = pdf_size
    pdf_ratio = pdf_width / pdf_height
    strategy = "diff"
    if strategy == "diff":
        # 根据宽高比设置PPT尺寸
        if pdf_ratio > 1.65:
            presentation.SlideSize.Type = SlideSizeType.Screen16x9
        elif pdf_ratio > 1.45:
            presentation.SlideSize.Type = SlideSizeType.Screen16x10
        elif pdf_ratio > 1.0:
            presentation.SlideSize.Type = SlideSizeType.Screen4x3
        else:
            ppt_height = 720
            ppt_width = ppt_height * pdf_ratio
            presentation.SlideSize.Type = SlideSizeType.Custom
            presentation.SlideSize.Size = SizeF(float(ppt_width),
                                                float(ppt_height))
    else:
        presentation.SlideSize.Type = SlideSizeType.Custom
        presentation.SlideSize.Size = SizeF(float(pdf_width),
                                            float(pdf_height))

    ppt_width = presentation.SlideSize.Size.Width
    ppt_height = presentation.SlideSize.Size.Height
    print(f"PPT Size: {ppt_width} x {ppt_height}")

    return presentation, ppt_width, ppt_height


# ============================================================================
# 文本块处理函数
# ============================================================================


def should_skip_text_block(label, content):
    """
    判断是否应该跳过该文本块
    
    Args:
        label: 文本块标签
        content: 文本内容
        
    Returns:
        bool: True表示跳过，False表示处理
    """
    # 只处理特定类型的文本标签
    valid_labels = [
        'text', 'title', 'header', 'footer', 'reference', 'paragraph_title',
        'algorithm'
    ]
    if label not in valid_labels:
        return True

    # 跳过页脚中的水印信息
    if label == 'footer' and "notebooklm" in content.lower():
        return True

    # 跳过空内容
    if not content.strip():
        return True

    return False


def create_text_shape(slide,
                      content,
                      label,
                      bbox,
                      scale,
                      ppt_width,
                      ppt_height,
                      font_size,
                      font_name,
                      delta_y=2):
    """
    在幻灯片上创建文本形状
    
    Args:
        slide: 幻灯片对象
        content: 文本内容
        label: 文本块标签
        bbox: 边界框 [x1, y1, x2, y2]
        scale: 缩放比例
        ppt_width: PPT宽度
        ppt_height: PPT高度
        font_size: 字体大小
        font_name: 字体名称
        delta_y: Y轴偏移量
        
    Returns:
        文本形状对象
    """
    bx1, by1, bx2, by2 = bbox

    # 坐标转换 (应用对齐偏移量 delta_y)
    left = bx1 * scale
    top = by1 * scale + delta_y
    right = bx2 * scale
    bottom = by2 * scale + delta_y
    width = right - left
    height = bottom - top


    if label =='paragraph_title':
        print(content)
        alignment = TextAlignmentType.Left
        h_padding = 5
        v_padding = 5
    else:
        alignment = TextAlignmentType.Center
        h_padding = 15
        v_padding = 5

    # 适当留白
    rect = RectangleF.FromLTRB(max(0, left ),
                               max(0, top - v_padding),
                               min(ppt_width, right + h_padding + h_padding),
                               min(ppt_height, bottom + v_padding))

    # 创建文本框

    text_shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect)
    text_shape.Name = f"Block_{label}"
    text_shape.TextFrame.Text = content
    text_shape.TextFrame.FitTextToShape = True

    text_shape.TextFrame.MarginLeft = 0
    text_shape.TextFrame.MarginRight = 0
    text_shape.TextFrame.MarginTop = 0
    text_shape.TextFrame.MarginBottom = 0

    text_shape.Line.FillType = FillFormatType.none
    text_shape.Fill.FillType = FillFormatType.none
    
    # 设置文本格式
    for paragraph in text_shape.TextFrame.Paragraphs:
        paragraph.Alignment = alignment
        for text_range in paragraph.TextRanges:
            text_range.LatinFont = TextFont(font_name)
            text_range.FontHeight = font_size
            text_range.Fill.FillType = FillFormatType.Solid
            text_range.Fill.SolidColor.Color = Color.FromArgb(255, 0, 0, 0)

    return text_shape


def process_text_blocks(slide,
                        parsing_res_list,
                        ocr_boxes,
                        scale,
                        ppt_width,
                        ppt_height,
                        font_name="Calibri"):
    """
    处理页面上的所有文本块
    
    Args:
        slide: 幻灯片对象
        parsing_res_list: 解析结果列表
        ocr_boxes: OCR文本框列表
        scale: 缩放比例
        ppt_width: PPT宽度
        ppt_height: PPT高度
        font_name: 字体名称
    """
    for item in parsing_res_list:
        label = item.get('block_label', 'unknown')
        content = item.get('block_content', '')
        bbox = item.get('block_bbox')

        if not bbox:
            continue

        # 判断是否跳过该文本块
        if should_skip_text_block(label, content):
            continue

        # 计算行数和字体大小
        line_count = get_line_count(bbox, ocr_boxes)
        is_multiline = line_count > 1

        bx1, by1, bx2, by2 = bbox
        width = (bx2 - bx1) * scale
        height = (by2 - by1) * scale

        font_size = calculate_font_size(height,
                                        is_multiline=is_multiline,
                                        line_count=line_count)

        # 创建文本形状
        create_text_shape(slide, content, label, bbox, scale, ppt_width,
                          ppt_height, font_size, font_name)


# ============================================================================
# 图片和背景处理函数
# ============================================================================


def expand_bbox(bbox, expand_px, size):
    """
    扩展边界框
    
    Args:
        bbox: 边界框 [x1, y1, x2, y2]
        expand_px: 扩展像素数
        img_width: 图片宽度
        img_height: 图片高度
        
    Returns:
        list: 扩展后的边界框
    """
    width, height = size
    x1, y1, x2, y2 = bbox
    x1 = max(0, x1 - expand_px)
    y1 = max(0, y1 - expand_px)
    x2 = min(width, x2 + expand_px)
    y2 = min(height, y2 + expand_px)
    return [x1, y1, x2, y2]


def scale_bbox(bbox, s, make_int=True):
    """
    按比例缩放并四舍五入边界框坐标
    
    Args:
        bbox: 边界框 [x1, y1, x2, y2]
        scale: 缩放比例
        
    Returns:
        list: 缩放并四舍五入后的边界框
    """
    if make_int:
        l, t, r, b = bbox
        # 左上向下取整，右下向上取整
        return [int(l * s), int(t * s), int(np.ceil(r * s)), int(np.ceil(b * s))]
    else:
        return [coord * s for coord in bbox]


def extract_foreground_element(slide, item, index, image_cv, img_scale, scale,
                               pdf_size, png_dir, page_idx):
    """
    提取前景元素(图片、表格、图表)并添加到幻灯片
    
    Args:
        slide: 幻灯片对象
        item: 元素信息
        index: 元素索引
        image_cv: 图片数组
        img_scale: 图片缩放比例
        scale: PPT缩放比例
        png_dir: PNG输出目录
        page_idx: 页面索引
        
    Returns:
        bool: 是否成功提取
    """
    label = item.get('block_label')
    bbox = item.get('block_bbox')

    if not bbox:
        return False

    expanded_bbox = expand_bbox(bbox, expand_px=2, size=pdf_size)

    # 原始图上的坐标用于裁剪 (对齐偏移量)
    l_img, t_img, r_img, b_img = scale_bbox(expanded_bbox, img_scale)

    if r_img <= l_img or b_img <= t_img:
        print("裁剪区域无效，跳过")
        return False

    # 裁剪并保存
    crop = image_cv[t_img:b_img, l_img:r_img]
    crop_name = f"page_{page_idx+1}_{label}_{index}.png"
    crop_path = png_dir / crop_name
    Image.fromarray(crop).save(crop_path)

    # 添加到幻灯片
    l_ppt, t_ppt, r_ppt, b_ppt = scale_bbox(expanded_bbox, scale, make_int=False)
    rect_item = RectangleF.FromLTRB(l_ppt, t_ppt, r_ppt, b_ppt)

    img_shape = slide.Shapes.AppendEmbedImageByPath(ShapeType.Rectangle,
                                                    str(crop_path), rect_item)
    img_shape.Line.FillType = FillFormatType.none
    img_shape.ZOrderPosition = 0  # 设为底层形状

    return True


def erase_region(image_cv, bbox, img_scale, pdf_size):
    """
    擦除图片中的指定区域
    
    Args:
        image_cv: 图片数组
        bbox: 边界框
        img_scale: 缩放比例
        pdf_size: PDF尺寸 (宽, 高)
        
    Returns:
        bool: 是否成功擦除
    """
    expanded_bbox = expand_bbox(bbox, expand_px=2, size=pdf_size)
    l, t, r, b = scale_bbox(expanded_bbox, img_scale, make_int=True)

    if r <= l or b <= t:
        print("擦除区域无效，跳过")
        return False

    # 计算填充颜色并擦除
    _, fill_color = compute_edge_diversity_numpy(image_cv,
                                                 l,
                                                 t,
                                                 r,
                                                 b,
                                                 tolerance=20)
    image_cv[t:b, l:r] = fill_color

    return True


def process_slide_background(slide, presentation, parsing_res_list, png_file,
                             pdf_size, scale, png_dir, page_idx):
    """
    处理幻灯片背景（提取前景元素、擦除已处理区域、设置背景）
    
    Args:
        slide: 幻灯片对象
        presentation: 演示文稿对象
        parsing_res_list: 解析结果列表
        png_file: PNG文件路径
        pdf_size: PDF尺寸 (宽, 高)
        scale: 缩放比例
        png_dir: PNG输出目录
        page_idx: 页面索引
    """
    if not png_file.exists():
        return

    # 加载图片
    pdf_w, pdf_h = pdf_size
    img = Image.open(png_file)
    img = img.resize(pdf_size, Image.LANCZOS)
    image_cv = np.array(img)
    image_h, image_w = image_cv.shape[:2]
    img_scale = image_w / pdf_w
    ppt2img_scale = scale / img_scale

    # 1. 提取前景图 (图片、表格、图表)
    for i, item in enumerate(parsing_res_list):
        label = item.get('block_label')
        if label in ['image', 'table', 'chart']:
            extract_foreground_element(slide, item, i, image_cv, img_scale,
                                       scale, pdf_size, png_dir, page_idx)

    # 2. 擦除已转换为文本框或独立图片的区域
    erasable_labels = [
        'text', 'title', 'header', 'footer', 'reference', 'paragraph_title',
        'image', 'table', 'algorithm', 'chart'
    ]

    for item in parsing_res_list:
        label = item.get('block_label')
        if label in erasable_labels:
            bbox = item.get('block_bbox')
            if bbox:
                erase_region(image_cv, bbox, img_scale, pdf_size)
                

    # 3. 保存处理后的图片
    processed_png = png_dir / f"page_{page_idx+1}_paddle_processed.png"
    Image.fromarray(image_cv).save(processed_png)

    # 4. 设置幻灯片背景
    slide.SlideBackground.Type = BackgroundType.Custom
    slide.SlideBackground.Fill.FillType = FillFormatType.Picture

    stream = Stream(str(processed_png))
    image_data = presentation.Images.AppendStream(stream)
    slide.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image_data
    slide.SlideBackground.Fill.PictureFill.FillType = PictureFillType.Stretch


def get_pdf_size_from_data(data):
    """
    从布局结果中获取PDF尺寸
    
    Args:
        layout_results: 布局解析结果列表
        
    Returns:
        tuple: (pdf_width, pdf_height)
    """
    layout_results = data.get('layoutParsingResults', [])
    first_page_layout = layout_results[0]['prunedResult']
    pdf_w = first_page_layout['width']
    pdf_h = first_page_layout['height']
    return (pdf_w, pdf_h)

def update_data_size(data, width, height=None):
    """
    设置数据中的宽高信息
    
    Args:
        data: PaddleOCR JSON数据
        width: 新的宽度
        height: 新的高度（可选，如果为None则只设置宽度）
        
    Returns:
        dict: 更新后的数据
    """
    width = int(width)
    if height is not None:
        height = int(height)
    layout_results = data.get('layoutParsingResults', [])
    
    # 更新布局结果中的宽高
    for page_layout in layout_results:
        pruned_result = page_layout.get('prunedResult', {})
        pruned_result['width'] = width
        if height is not None:
            pruned_result['height'] = height
    
    # 更新dataInfo中的宽高
    data_info = data.get('dataInfo', {})
    if data_info:
        data_info['width'] = width
        if height is not None:
            data_info['height'] = height
    
    # 更新pages中的宽高
    pages_info = data_info.get('pages', [])
    if pages_info:
        for page in pages_info:
            page['width'] = width
            if height is not None:
                page['height'] = height
    
    return data

def make_data_wide_screen(data):
    """
    将数据调整为16:9宽屏比例
    
    Args:
        data: PaddleOCR JSON数据
        
    Returns:
        dict: 调整后的数据
    """
    data = copy.deepcopy(data)
    
    # 获取原始PDF尺寸
    layout_results = data.get('layoutParsingResults', [])
    if not layout_results:
        return data
    
    first_page_layout = layout_results[0]['prunedResult']
    pdf_w = first_page_layout['width']
    pdf_h = first_page_layout['height']
    
    # 计算16:9宽屏的目标宽度
    target_width = round(pdf_h * 16 / 9)
    
    if target_width == pdf_w:
        print(f"✓ 已是16:9宽屏，无需调整")
        return data
    
    if target_width > pdf_w:
        # 需要扩展宽度 - 计算左右偏移量
        offset_x = (target_width - pdf_w) // 2
        
        print(f"✓ 扩展宽度: {pdf_w} -> {target_width} (左右各偏移 {offset_x})")
        
        # 设置新的宽度
        data = update_data_size(data, target_width)
        layout_results = data.get('layoutParsingResults', [])
        for page_layout in layout_results:
            pruned_result = page_layout.get('prunedResult', {})
            parsing_res_list = pruned_result.get('parsing_res_list', [])
            for item in parsing_res_list:
                bbox = item.get('block_bbox')
                if bbox:
                    # 横坐标右移
                    item['block_bbox'] = [bbox[0] + offset_x, bbox[1], 
                                            bbox[2] + offset_x, bbox[3]]
        
        # 调整OCR结果
        ocr_results = data.get('ocrResults', [])
        for page_ocr in ocr_results:
            pruned_result = page_ocr.get('prunedResult', {})
            rec_boxes = pruned_result.get('rec_boxes', [])
            
            for bbox in rec_boxes:
                if bbox and len(bbox) >= 4:
                    # 横坐标右移
                    bbox[0] += offset_x
                    bbox[2] += offset_x
    
    elif target_width < pdf_w:
        # 需要裁剪宽度 - 计算左边界
        left = (pdf_w - target_width) // 2
        right = left + target_width
        
        print(f"✓ 裁剪宽度: {pdf_w} -> {target_width} (保留中间部分 {left}-{right})")
        
        # 设置新的宽度
        data = update_data_size(data, target_width)
        layout_results = data.get('layoutParsingResults', [])
        
        # 调整所有坐标并过滤超出范围的元素
        for page_layout in layout_results:
            pruned_result = page_layout.get('prunedResult', {})
            pruned_result['width'] = target_width
            
            parsing_res_list = pruned_result.get('parsing_res_list', [])
            filtered_list = []
            
            for item in parsing_res_list:
                bbox = item.get('block_bbox')
                if bbox:
                    x1, y1, x2, y2 = bbox
                    # 检查是否在裁剪范围内
                    if x2 > left and x1 < right:
                        # 调整坐标并裁剪到边界
                        new_x1 = max(0, x1 - left)
                        new_x2 = min(target_width, x2 - left)
                        item['block_bbox'] = [new_x1, y1, new_x2, y2]
                        filtered_list.append(item)
            
            pruned_result['parsing_res_list'] = filtered_list
        
        # 调整OCR结果
        ocr_results = data.get('ocrResults', [])
        for page_ocr in ocr_results:
            pruned_result = page_ocr.get('prunedResult', {})
            rec_boxes = pruned_result.get('rec_boxes', [])
            
            filtered_boxes = []
            for bbox in rec_boxes:
                if bbox and len(bbox) >= 4:
                    x1, y1, x2, y2 = bbox[0], bbox[1], bbox[2], bbox[3]
                    # 检查是否在裁剪范围内
                    if x2 > left and x1 < right:
                        # 调整坐标
                        bbox[0] = max(0, x1 - left)
                        bbox[2] = min(target_width, x2 - left)
                        filtered_boxes.append(bbox)
            
            pruned_result['rec_boxes'] = filtered_boxes
    
    return data



def resize_data(data, pdf_size, ppt_size):
    """
    调整数据中的坐标以适应PPT尺寸
    
    Args:
        data: PaddleOCR JSON数据
        pdf_size: PDF尺寸 (pdf_width, pdf_height)
        ppt_size: PPT尺寸 (ppt_width, ppt_height)
        
    Returns:
        dict: 调整后的数据
    """
    pdf_w, pdf_h = pdf_size
    ppt_w, ppt_h = ppt_size
    
    scale_x = ppt_w / pdf_w
    scale_y = ppt_h / pdf_h
    print(f"Resize Data: scale_x={scale_x}, scale_y={scale_y}")

    assert abs(scale_x - scale_y) < 1e-2, "X和Y缩放比例不一致"
    scale = scale_x
    data = copy.deepcopy(data)
    
    # 同步更新页面尺寸信息
    data = update_data_size(data, ppt_w, ppt_h)
    
    layout_results = data.get('layoutParsingResults', [])
    ocr_results = data.get('ocrResults', [])
    
    # 调整布局结果中的坐标
    for page_layout in layout_results:
        pruned_result = page_layout.get('prunedResult', {})
        parsing_res_list = pruned_result.get('parsing_res_list', [])
        
        for item in parsing_res_list:
            bbox = item.get('block_bbox')
            if bbox:
                # 调整边界框坐标
                item['block_bbox'] = scale_bbox(bbox, scale, make_int=True)
    
    # 调整OCR结果中的坐标
    for page_ocr in ocr_results:
        pruned_result = page_ocr.get('prunedResult', {})
        rec_boxes = pruned_result.get('rec_boxes', [])
        
        for bbox in rec_boxes:
            if bbox and len(bbox) >= 4:
                # 调整OCR框坐标
                new_bbox = scale_bbox(bbox, scale, make_int=True)
                bbox[0], bbox[1], bbox[2], bbox[3] = new_bbox
    
    return data



# ============================================================================
# 主要处理函数
# ============================================================================


def create_ppt_from_paddle_json(json_file,
                                pdf_file,
                                output_dir,
                                out_ppt_name=None,
                                dpi=150,
                                inpaint=True,
                                inpaint_method='background_smooth'):
    """
    从 PaddleOCR JSON 直接创建 PPT
    
    Args:
        json_file: PaddleOCR JSON文件路径
        pdf_file: 原始PDF文件路径
        output_dir: 输出目录
        out_ppt_name: 输出PPT文件名
        dpi: 图片清晰度
        inpaint: 是否进行图像修复
        inpaint_method: 图像修复方法
    """
    # 验证输入文件
    if not os.path.exists(json_file):
        print(f"错误: JSON 文件 {json_file} 不存在")
        return

    if not os.path.exists(pdf_file):
        print(f"错误: PDF 文件 {pdf_file} 不存在")
        return

    # 准备输出目录
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True, parents=True)
    png_dir = output_dir / "png"
    png_dir.mkdir(exist_ok=True)

    # 步骤 1: 将 PDF 转换为 PNG
    print("=" * 60)
    print("步骤 1: 将 PDF 转换为 PNG 图片")
    print("=" * 60)
    png_names = pdf_to_png(pdf_file,
                           png_dir,
                           dpi=dpi,
                           inpaint=inpaint,
                           inpaint_method=inpaint_method,
                           force_regenerate=True, 
                           make_wide_screen=True)
    # 重新利用生成的 PNG 文件列表，生成PDF
    
    png_files = [png_dir / name for name in png_names]

    # 步骤 2: 读取 JSON 文件
    print("\n" + "=" * 60)
    print("步骤 2: 读取 PaddleOCR JSON 文件")
    print("=" * 60)
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 步骤 2.5: 调整为宽屏比例
    print("\n" + "=" * 60)
    print("步骤 2.5: 调整数据为16:9宽屏比例")
    print("=" * 60)
    data = make_data_wide_screen(data)

    # 步骤 3: 创建 PPT
    print("\n" + "=" * 60)
    print("步骤 3: 从 PaddleOCR JSON 创建最终 PPT")
    print("=" * 60)

    # 获取PDF尺寸
    pdf_size = get_pdf_size_from_data(data)
    print(f"PDF Size from data: {pdf_size[0]} x {pdf_size[1]}")

    # 设置PPT并调整数据坐标
    presentation, ppt_width, ppt_height = setup_presentation(pdf_size)
    data = resize_data(data, pdf_size, (ppt_width, ppt_height))
    pdf_size = get_pdf_size_from_data(data)
    scale = ppt_width / pdf_size[0]
    assert abs(scale - 1.0) < 1e-2, "resize_data后scale应≈1"
    scale = 1.0
    
    # 更新数据引用
    layout_results = data.get('layoutParsingResults', [])
    ocr_results = data.get('ocrResults', [])

    font_name = "Calibri"

    # 处理每一页
    for page_idx in range(len(layout_results)):
        print(f"处理第 {page_idx+1}/{len(layout_results)} 页...")
        slide = presentation.Slides.Append()

        page_layout = layout_results[page_idx]['prunedResult']
        page_ocr = ocr_results[page_idx]['prunedResult']

        parsing_res_list = page_layout.get('parsing_res_list', [])
        ocr_boxes = page_ocr.get('rec_boxes', [])

        # 处理文本块
        process_text_blocks(slide, parsing_res_list, ocr_boxes, scale,
                            ppt_width, ppt_height, font_name)

        # 处理背景图片（包括前景元素提取和区域擦除）
        if page_idx < len(png_files):
            process_slide_background(slide, presentation, parsing_res_list,
                                     png_files[page_idx], pdf_size, scale,
                                     png_dir, page_idx)

    # 保存并清理PPT
    if out_ppt_name is None:
        out_ppt_name = os.path.basename(pdf_file).replace('.pdf', '.pptx')
    final_ppt_file = output_dir / out_ppt_name
    presentation.SaveToFile(str(final_ppt_file), FileFormat.Pptx2019)
    clean_ppt(str(final_ppt_file), str(final_ppt_file))
    print(f"\n完成! 输出文件: {final_ppt_file}")


# ============================================================================
# 命令行入口
# ============================================================================


def main():
    """命令行入口函数"""
    import sys
    sys.stdout.reconfigure(encoding='utf-8')

    parser = argparse.ArgumentParser(
        description="从 PaddleOCR JSON 直接创建 PPT",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
python create_with_paddle_reorg.py result.json input.pdf -o output --dpi 150
        """)

    parser.add_argument("json_file", help="PaddleOCR JSON文件路径 (result.json)")
    parser.add_argument("pdf_file", help="原始PDF文件路径")
    parser.add_argument("--workspace",
                        default="output",
                        type=str,
                        help="工作目录 (默认: output)")
    parser.add_argument('--name', type=str, default=None)
    parser.add_argument("--dpi", type=int, default=150, help="图片清晰度 (默认: 150)")

    args = parser.parse_args()

    workspace = Path(args.workspace)
    out_dir = workspace / os.path.basename(args.pdf_file).replace('.pdf', '')
    out_dir.mkdir(exist_ok=True, parents=True)

    create_ppt_from_paddle_json(args.json_file,
                                args.pdf_file,
                                str(out_dir),
                                out_ppt_name=args.name,
                                dpi=args.dpi)


if __name__ == "__main__":
    main()
