import os
import sys
from pathlib import Path
import win32com.client
import fitz  # PyMuPDF
from PIL import Image

# 确保可以导入项目中的模块
sys.path.append(str(Path(__file__).parent.parent))
from notebooklm2ppt.pdf2png import pdf_to_png

def pptx_to_pdf(pptx_path, pdf_output_path):
    """
    将PPTX文件转换为PDF文件
    
    Args:
        pptx_path: PPTX文件路径
        pdf_output_path: 输出PDF文件路径
    """
    print(f"正在将PPTX转换为PDF: {pptx_path}")
    
    # 创建PowerPoint应用程序对象
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    
    # 打开PPTX文件
    presentation = powerpoint.Presentations.Open(pptx_path)
    
    # 保存为PDF
    presentation.SaveAs(pdf_output_path, 32)  # 32 = PDF格式
    
    # 关闭演示文稿
    presentation.Close()
    
    # 退出PowerPoint
    powerpoint.Quit()
    
    print(f"✓ 已保存PDF: {pdf_output_path}")

def compare_pdf_and_pptx(pdf_path, pptx1_path, pptx2_path, output_dir=None, dpi=150):
    """
    比较PDF文件和两个由其转换的PPTX文件
    
    Args:
        pdf_path: 源PDF文件路径
        pptx1_path: 第一个PPTX文件路径
        pptx2_path: 第二个PPTX文件路径
        output_dir: 输出目录，默认为PDF同目录的compare_results文件夹
        dpi: 图片清晰度，默认150
    """
    # 确定输出目录
    if output_dir is None:
        pdf_name = Path(pdf_path).stem
        output_dir = Path(pdf_path).parent / "compare_results"
    else:
        output_dir = Path(output_dir)
    
    # 创建输出目录
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 转换PPTX为PDF
    pptx1_pdf = output_dir / f"{Path(pptx1_path).stem}.pdf"
    pptx2_pdf = output_dir / f"{Path(pptx2_path).stem}.pdf"
    
    pptx_to_pdf(pptx1_path, str(pptx1_pdf))
    pptx_to_pdf(pptx2_path, str(pptx2_pdf))
    
    # 转换PDF为JPG（使用同一个目录，通过文件名后缀区分）
    jpg_dir = output_dir / "jpgs"
    
    print("\n正在将源PDF转换为JPG...")
    # 为了区分不同来源的图片，我们先转换到临时目录，然后重命名到统一目录
    temp_pdf_dir = output_dir / "temp_pdf"
    pdf_jpgs = pdf_to_png(pdf_path, temp_pdf_dir, dpi=dpi, force_regenerate=True)
    
    print("\n正在将第一个PPTX转换的PDF转换为JPG...")
    temp_pptx1_dir = output_dir / "temp_pptx1"
    pptx1_jpgs = pdf_to_png(str(pptx1_pdf), temp_pptx1_dir, dpi=dpi, force_regenerate=True)
    
    print("\n正在将第二个PPTX转换的PDF转换为JPG...")
    temp_pptx2_dir = output_dir / "temp_pptx2"
    pptx2_jpgs = pdf_to_png(str(pptx2_pdf), temp_pptx2_dir, dpi=dpi, force_regenerate=True)
    
    # 创建统一的JPG目录
    jpg_dir.mkdir(exist_ok=True)
    
    # 重命名并移动所有JPG文件到统一目录
    print("\n正在整理JPG文件到统一目录...")
    
    # 先确定统一的目标尺寸（使用源PDF的图片尺寸作为标准）
    target_width = 0
    target_height = 0
    
    # 获取源PDF图片的尺寸作为标准
    for jpg_name in pdf_jpgs:
        src_path = temp_pdf_dir / jpg_name
        if src_path.exists():
            from PIL import Image
            img = Image.open(src_path)
            target_width = max(target_width, img.width)
            target_height = max(target_height, img.height)
    
    # 如果没有源PDF图片，使用默认尺寸
    if target_width == 0 or target_height == 0:
        target_width = 800
        target_height = 600
    
    print(f"\n使用统一尺寸: {target_width} x {target_height}")
    
    # 移动源PDF的JPG文件
    for jpg_name in pdf_jpgs:
        src_path = temp_pdf_dir / jpg_name
        # 使用数字后缀确保源PDF转换的图片排在前面：page_0001_0.jpg
        base_name = jpg_name.rsplit('.', 1)[0]
        # 强制使用jpg扩展名
        new_name = f"{base_name}_0_source.jpg"
        dst_path = jpg_dir / new_name
        if src_path.exists():
            # 先转换为jpg格式并调整尺寸
            from PIL import Image
            img = Image.open(src_path)
            # 调整到统一尺寸
            img = img.resize((target_width, target_height))
            img.save(dst_path, 'JPEG', quality=95)
            print(f"  ✓ 已转换并调整尺寸: {dst_path.name}")
    
    # 移动第一个PPTX的JPG文件
    for jpg_name in pptx1_jpgs:
        src_path = temp_pptx1_dir / jpg_name
        # 使用数字后缀：page_0001_1.jpg
        base_name = jpg_name.rsplit('.', 1)[0]
        # 强制使用jpg扩展名
        new_name = f"{base_name}_1_converted.jpg"
        dst_path = jpg_dir / new_name
        if src_path.exists():
            # 先转换为jpg格式并调整尺寸
            from PIL import Image
            img = Image.open(src_path)
            # 调整到统一尺寸
            img = img.resize((target_width, target_height))
            img.save(dst_path, 'JPEG', quality=95)
            print(f"  ✓ 已转换并调整尺寸: {dst_path.name}")
    
    # 移动第二个PPTX的JPG文件
    for jpg_name in pptx2_jpgs:
        src_path = temp_pptx2_dir / jpg_name
        # 使用数字后缀：page_0001_2.jpg
        base_name = jpg_name.rsplit('.', 1)[0]
        # 强制使用jpg扩展名
        new_name = f"{base_name}_2_converted.jpg"
        dst_path = jpg_dir / new_name
        if src_path.exists():
            # 先转换为jpg格式并调整尺寸
            from PIL import Image
            img = Image.open(src_path)
            # 调整到统一尺寸
            img = img.resize((target_width, target_height))
            img.save(dst_path, 'JPEG', quality=95)
            print(f"  ✓ 已转换并调整尺寸: {dst_path.name}")
    
    # 清理临时目录
    import shutil
    for temp_dir in [temp_pdf_dir, temp_pptx1_dir, temp_pptx2_dir]:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
            print(f"  ✓ 已清理临时目录: {temp_dir.name}")
    
    # 确定最大页数
    max_pages = max(len(pdf_jpgs), len(pptx1_jpgs), len(pptx2_jpgs))
    print(f"\n最大页数: {max_pages}")
    
    # 横向拼接图片
    print("\n正在横向拼接图片...")
    combined_dir = output_dir / "combined"
    combined_dir.mkdir(exist_ok=True)
    
    for i in range(max_pages):
        page_num = i + 1
        print(f"  处理第 {page_num} 页...")
        
        # 获取对应页面的JPG文件路径（从统一目录中获取带有数字后缀的文件）
        pdf_jpg_path = jpg_dir / f"page_{page_num:04d}_0_source.jpg"
        pptx1_jpg_path = jpg_dir / f"page_{page_num:04d}_1_converted.jpg"
        pptx2_jpg_path = jpg_dir / f"page_{page_num:04d}_2_converted.jpg"
        
        # 加载图片
        images = []
        
        # 确定统一的图片尺寸
        target_width = 0
        target_height = 0
        
        # 先获取所有图片的尺寸，确定最大尺寸
        img_paths = [pdf_jpg_path, pptx1_jpg_path, pptx2_jpg_path]
        for path in img_paths:
            if path.exists():
                img = Image.open(path)
                target_width = max(target_width, img.width)
                target_height = max(target_height, img.height)
        
        # 如果没有找到任何图片，使用默认尺寸
        if target_width == 0 or target_height == 0:
            target_width = 800
            target_height = 600
        
        # 加载并调整所有图片到统一尺寸
        for path in img_paths:
            if path.exists():
                img = Image.open(path)
                # 调整图片到统一尺寸
                img = img.resize((target_width, target_height))
            else:
                # 创建空白图片
                img = Image.new('RGB', (target_width, target_height), color='white')
            images.append(img)
        
        # 计算拼接后的图片尺寸
        total_width = target_width * 3
        max_height = target_height
        
        # 创建拼接后的图片
        combined = Image.new('RGB', (total_width, max_height), color='white')
        
        # 拼接图片
        x_offset = 0
        for img in images:
            combined.paste(img, (x_offset, 0))
            x_offset += target_width
        
        # 保存拼接后的图片
        combined_path = combined_dir / f"combined_page_{page_num:04d}.jpg"
        combined.save(combined_path, 'JPEG', quality=95)
        print(f"  ✓ 已保存: {combined_path}")
    
    print(f"\n✅ 比较完成！")
    print(f"📁 输出目录: {output_dir}")
    print(f"📁 拼接后的图片目录: {combined_dir}")

def main():
    """
    主函数
    """
    if len(sys.argv) != 4:
        print("用法: python compare_result.py <pdf文件> <pptx文件1> <pptx文件2>")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    pptx1_path = sys.argv[2]
    pptx2_path = sys.argv[3]
    
    # 检查文件是否存在
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在: {pdf_path}")
        sys.exit(1)
    
    if not os.path.exists(pptx1_path):
        print(f"错误: PPTX文件1不存在: {pptx1_path}")
        sys.exit(1)
    
    if not os.path.exists(pptx2_path):
        print(f"错误: PPTX文件2不存在: {pptx2_path}")
        sys.exit(1)
    
    # 执行比较
    compare_pdf_and_pptx(pdf_path, pptx1_path, pptx2_path)

if __name__ == "__main__":
    main()
