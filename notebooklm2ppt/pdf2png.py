import fitz  # PyMuPDF
import os
from pathlib import Path
from .utils.image_inpainter import inpaint_image, INPAINT_METHODS
from PIL import Image

def pdf_to_png(pdf_path, output_dir=None, dpi=150, inpaint=False, pages=None, inpaint_method='background_smooth', force_regenerate=False, make_wide_screen=False):
    """
    将 PDF 文件转换为多个 PNG 图片
    
    参数:
        pdf_path: PDF 文件路径
        output_dir: 输出目录，默认为 PDF 同目录的 pdf_name_pngs 文件夹
        dpi: 分辨率，默认 150
        inpaint: 是否进行图像修复
        pages: 要处理的页码范围
        inpaint_method: 修复方法，可选值: background_smooth, edge_mean_smooth, background, onion, griddata, skimage
        force_regenerate: 是否强制重新生成所有 PNG（默认 False，复用已存在的 PNG）
        make_wide_screen: 是否变为宽屏图片，适应16:9 PPT页面
    """
    # 打开 PDF 文件
    pdf_doc = fitz.open(pdf_path)
    
    # 确定输出目录
    if output_dir is None:
        pdf_name = Path(pdf_path).stem  # 获取 PDF 文件名（不含扩展名）
        output_dir = Path(pdf_path).parent / f"{pdf_name}_pngs"
    else:
        output_dir = Path(output_dir)
    
    # 创建输出目录
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 转换因子：DPI / 72（默认屏幕 DPI）
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    
    # 遍历每一页
    page_count = len(pdf_doc)  # 在关闭文档前获取页数
    # 规范 pages 参数为集合，方便判断
    pages_set = None
    if pages is not None:
        pages_set = set(pages)

    png_names = []
    for page_num, page in enumerate(pdf_doc, 1):
        # 如果指定了 pages，则跳过不在范围内的页码
        if pages_set is not None and page_num not in pages_set:
            continue
        # 渲染页面为图片
        pix = page.get_pixmap(matrix=mat, alpha=False)
        
        # 保存为 PNG
        output_path = output_dir / f"page_{page_num:04d}.png"
        
        png_names.append(output_path.name)

        if not force_regenerate and os.path.exists(output_path):
            print(f"跳过已存在的文件: {output_path}")
            continue
        pix.save(output_path)
        print(f"✓ 已保存: {output_path}")
        if inpaint:
            inpaint_image(str(output_path), str(output_path), inpaint_method=inpaint_method)
            print(f"✓ 已修复: {output_path}")

        if make_wide_screen:
            # 变为16:9宽屏图片
            img = Image.open(output_path)
            width, height = img.size
            target_width = round(height * 16 / 9)
            if target_width > width:
                # 需要扩展宽度
                new_img = Image.new("RGB", (target_width, height), (255, 255, 255))
                new_img.paste(img, ((target_width - width) // 2, 0))
                new_img.save(output_path)
                print(f"✓ 已调整为宽屏: {output_path}")
            elif target_width < width:
                # 需要裁剪宽度
                left = (width - target_width) // 2
                right = left + target_width
                cropped_img = img.crop((left, 0, right, height))
                cropped_img.save(output_path)
                print(f"✓ 已裁剪为宽屏: {output_path}")
            else:
                print(f"✓ 已是宽屏，无需调整: {output_path}")

    pdf_doc.close()
    print(f"\n完成! 共转换 {page_count} 页，输出目录: {output_dir}")
    return png_names

def pngs2pdf(png_files, output_pdf):
    """
    将多个 PNG 图片合并为一个 PDF 文件
    
    参数:
        png_files: PNG 文件路径列表
        output_pdf: 输出 PDF 文件路径
    """
    if not png_files:
        print("没有提供 PNG 文件列表，无法生成 PDF。")
        return

    image_list = []
    for png_file in png_files:
        img = Image.open(png_file).convert("RGB")
        image_list.append(img)

    # 保存为 PDF
    first_image = image_list[0]
    rest_images = image_list[1:]
    first_image.save(output_pdf, save_all=True, append_images=rest_images)
    print(f"✓ 已生成 PDF: {output_pdf}")

if __name__ == "__main__":
    # 使用示例
    pdf_file = r"examples\Floyd_算法的动态规划之魂.pdf"  # 修改为你的 PDF 文件路径
    
    if os.path.exists(pdf_file):
        out_dir = 'tmp_pngs'
        png_names = pdf_to_png(pdf_file, dpi=150, output_dir=out_dir, make_wide_screen=True)
        png_files = [os.path.join(out_dir, name) for name in png_names]
        output_pdf = "output_widescreen.pdf"
        pngs2pdf(png_files, output_pdf)
    else:
        print(f"错误: 文件 {pdf_file} 不存在")
