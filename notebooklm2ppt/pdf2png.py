import fitz  # PyMuPDF
import os
from pathlib import Path
from .utils.image_inpainter import inpaint_image

def pdf_to_png(pdf_path, output_dir=None, dpi=150, inpaint=False, pages=None):
    """
    将 PDF 文件转换为多个 PNG 图片
    
    参数:
        pdf_path: PDF 文件路径
        output_dir: 输出目录，默认为 PDF 同目录的 pdf_name_pngs 文件夹
        dpi: 分辨率，默认 150
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

        if os.path.exists(output_path):
            print(f"跳过已存在的文件: {output_path}")
            continue
        pix.save(output_path)
        print(f"✓ 已保存: {output_path}")
        if inpaint:
            inpaint_image(str(output_path), str(output_path))
            print(f"✓ 已修复: {output_path}")

        
            
    pdf_doc.close()
    print(f"\n完成! 共转换 {page_count} 页，输出目录: {output_dir}")
    return png_names

if __name__ == "__main__":
    # 使用示例
    pdf_file = "Hackathon_Architect_Playbook.pdf"  # 修改为你的 PDF 文件路径
    
    if os.path.exists(pdf_file):
        pdf_to_png(pdf_file, dpi=150)
    else:
        print(f"错误: 文件 {pdf_file} 不存在")
