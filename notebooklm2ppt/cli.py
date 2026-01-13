"""命令行界面：将 PDF 转换为可编辑 PowerPoint 演示文稿"""

import os
import time
import threading
import cv2
import shutil
import argparse
import sys
from pathlib import Path
from .pdf2png import pdf_to_png
from .utils.image_viewer import show_image_fullscreen
from .utils.screenshot_automation import take_fullscreen_snip, mouse, screen_height, screen_width, load_saved_done_offset
from .ppt_combiner import combine_ppt


def process_pdf_to_ppt(pdf_path, png_dir, ppt_dir, delay_between_images=2, inpaint=True, dpi=150, timeout=50, display_height=None, 
                    display_width=None, done_button_offset=None, capture_done_offset: bool = True, pages=None, update_offset_callback=None):
    """
    将 PDF 转换为 PNG 图片，然后对每张图片进行截图处理
    
            try:
        pdf_path: PDF 文件路径
            # 对第一页允许用户手动点击并捕获完成按钮偏移（如果未保存且启用）
            capture_offset = (idx == 1 and capture_done_offset)
            success, ppt_filename, computed_offset = take_fullscreen_snip(
                check_ppt_window=True,
                ppt_check_timeout=timeout,
                width=display_width,
                height=display_height,
                pc_manager_version=pc_manager_version,
                done_button_right_offset=done_button_offset,
                capture_done_offset_if_missing=capture_offset,
                force_capture=force_calibrate,
            )
            # 如果本次截屏捕获并保存了偏移，回调通知（GUI 可使用此回调更新显示）
            try:
                if computed_offset is not None and offset_saved_callback:
                    offset_saved_callback(int(computed_offset))
            except Exception:
                pass
        pc_manager_version: 电脑管家版本号；3.19及以上自动使用 190，低于3.19 使用 210
        done_button_offset: 完成按钮右侧偏移量，传入数字时优先使用，不传则按版本推断
    """
    # 1. 将 PDF 转换为 PNG 图片
    print("=" * 60)
    print("步骤 1: 将 PDF 转换为 PNG 图片")
    print("=" * 60)
    
    if not os.path.exists(pdf_path):
        print(f"错误: PDF 文件 {pdf_path} 不存在")
        return
    
    png_names = pdf_to_png(pdf_path, png_dir, dpi=dpi, inpaint=inpaint, pages=pages)
    
    # 创建ppt输出目录
    ppt_dir.mkdir(exist_ok=True, parents=True)
    print(f"PPT输出目录: {ppt_dir}")
    
    # 获取用户的下载文件夹路径
    downloads_folder = Path.home() / "Downloads"
    print(f"下载文件夹: {downloads_folder}")
    
    # 2. 获取所有 PNG 图片文件并排序
    png_files = [png_dir / name for name in png_names]
    
    if not png_files:
        print(f"错误: 在 {png_dir} 中没有找到 PNG 图片")
        return
    
    print("\n" + "=" * 60)
    print(f"步骤 2: 处理 {len(png_files)} 张 PNG 图片")
    print("=" * 60)
    
    # 设置显示窗口尺寸（如果未指定则使用屏幕尺寸）
    if display_height is None:
        display_height = screen_height
    if display_width is None:
        display_width = screen_width
    
    print(f"显示窗口尺寸: {display_width} x {display_height}")

    
    # 3. 对每张图片进行截图处理
    for idx, png_file in enumerate(png_files, 1):
        print(f"\n[{idx}/{len(png_files)}] 处理图片: {png_file.name}")
        
        stop_event = threading.Event()
        
        def _viewer():
            """在线程中显示图片"""
            show_image_fullscreen(str(png_file), display_height=display_height)
            # 维持 OpenCV 事件循环
            while not stop_event.is_set():
                cv2.waitKey(50)
            # 关闭窗口
            try:
                cv2.destroyAllWindows()
            except Exception:
                pass
        
        # 启动图片显示线程
        viewer_thread = threading.Thread(
            target=_viewer, 
            name=f"opencv_viewer_{idx}", 
            daemon=True
        )
        viewer_thread.start()
        
        # 等待窗口稳定
        time.sleep(3)
        
        try:
            # 执行全屏截图并检测PPT窗口
            # 对第一页允许用户手动点击并捕获完成按钮偏移（如果未保存或被强制要求）
            capture_offset = (idx == 1 and capture_done_offset)
            if capture_offset:
                done_button_offset = None  # 强制重新捕获偏移
            else:
                assert done_button_offset is not None, "必须提供完成按钮偏移量"
            success, ppt_filename, computed_offset = take_fullscreen_snip(
                check_ppt_window=True,
                ppt_check_timeout=timeout,
                width=display_width,
                height=display_height,
                done_button_right_offset=done_button_offset,
            )
            if success and computed_offset is not None:
                print(f"捕获到的完成按钮偏移: {computed_offset}")
                done_button_offset = computed_offset  # 更新为最新捕获的偏移
                if update_offset_callback:
                    update_offset_callback()


            if success and ppt_filename:
                print(f"✓ 图片 {png_file.name} 处理完成，PPT窗口已创建: {ppt_filename}")
                
                # 从下载文件夹查找并复制PPT文件
                if " - PowerPoint" in ppt_filename:
                    base_filename = ppt_filename.replace(" - PowerPoint", "").strip()
                else:
                    base_filename = ppt_filename.strip()
                
                if not base_filename.endswith(".pptx"):
                    search_filename = base_filename + ".pptx"
                else:
                    search_filename = base_filename
                
                ppt_source_path = downloads_folder / search_filename
                
                if not ppt_source_path.exists():
                    print(f"  未找到 {ppt_source_path}，尝试查找最近的.pptx文件...")
                    pptx_files = list(downloads_folder.glob("*.pptx"))
                    if pptx_files:
                        pptx_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
                        ppt_source_path = pptx_files[0]
                        print(f"  找到最近的PPT文件: {ppt_source_path.name}")
                
                if ppt_source_path.exists():
                    target_filename = png_file.stem + ".pptx"
                    target_path = ppt_dir / target_filename
                    
                    shutil.copy2(ppt_source_path, target_path)
                    print(f"  ✓ PPT文件已复制: {target_path}")
                    
                    try:
                        ppt_source_path.unlink()
                        print(f"  ✓ 已删除原文件: {ppt_source_path}")
                    except Exception as e:
                        print(f"  ⚠ 删除原文件失败: {e}")
                else:
                    print(f"  ⚠ 未在下载文件夹中找到PPT文件")
            elif success:
                print(f"✓ 图片 {png_file.name} 处理完成，但未获取到PPT文件名")
            else:
                print(f"⚠ 图片 {png_file.name} 已截图，但未检测到新的PPT窗口")
                close_button = (display_width - 35, display_height + 35)
                mouse.click(button='left', coords=close_button)
        except Exception as e:
            print(f"✗ 处理图片 {png_file.name} 时出错: {e}")
        finally:
            stop_event.set()
            viewer_thread.join(timeout=2)
        
        if idx < len(png_files):
            print(f"等待 {delay_between_images} 秒后处理下一张...")
            time.sleep(delay_between_images)
    
    print("\n" + "=" * 60)
    print(f"完成! 共处理 {len(png_files)} 张图片")
    print("=" * 60)
    return png_names


def main():
    # 如果没有参数，或者第一个参数是 --gui，则启动 GUI
    if len(sys.argv) == 1 or (len(sys.argv) > 1 and sys.argv[1] == "--gui"):
        from .gui import launch_gui
        launch_gui()
        return

    # 解析命令行参数
    parser = argparse.ArgumentParser(
        description='NotebookLM2PPT - 将 PDF 转换为可编辑 PowerPoint 演示文稿',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  notebooklm2ppt examples/demo.pdf                  # 转换指定PDF
  notebooklm2ppt examples/demo.pdf --no-inpaint     # 禁用图像修复（去水印）
  notebooklm2ppt examples/demo.pdf -d 3 -t 60       # 设置延迟和超时
  notebooklm2ppt -s 0.9 examples/demo.pdf           # 设置显示尺寸比例
        """
    )
    
    parser.add_argument(
        'pdf_file',
        help='PDF 文件路径'
    )
    
    parser.add_argument(
        '-d', '--delay',
        type=float,
        default=2,
        metavar='SECONDS',
        help='处理每张图片之间的延迟时间，单位秒（默认: 2）'
    )
    
    parser.add_argument(
        '-t', '--timeout',
        type=float,
        default=50,
        metavar='SECONDS',
        help='PPT 窗口检测超时时间，单位秒（默认: 50）'
    )
    
    parser.add_argument(
        '--inpaint-notebooklm',
        dest='inpaint',
        action='store_true',
        help='启用图像修复功能（去水印），只能去除notebooklm生成的水印'
    )
    
    parser.add_argument(
        '--no-inpaint',
        dest='inpaint',
        action='store_false',
        help='禁用图像修复功能'
    )

    parser.add_argument(
        '--dpi',
        type=int,
        default=150,
        metavar='DPI',
        help='PNG 输出分辨率，必须为150以启用图像修复（默认: 150）'
    )
    
    parser.set_defaults(inpaint=True)

    # 根据磁盘上是否存在已保存偏移决定默认是否进行首次校准
    try:
        _saved_offset = load_saved_done_offset()
        _default_capture = True if _saved_offset is None else False
    except Exception:
        _default_capture = True
    parser.set_defaults(capture_done_offset=_default_capture)
    
    parser.add_argument(
        '-s', '--size-ratio',
        type=float,
        default=0.8,
        metavar='RATIO',
        help='显示尺寸比例，1.0 表示填满屏幕（默认: 0.8），如果转换失败，可尝试调小此值重试'
    )
    
    parser.add_argument(
        '-o', '--output',
        metavar='DIR',
        help='输出目录（默认: workspace）'
    )

    parser.add_argument(
        '--pcmgr-version',
        dest='pc_manager_version',
        type=str,
        default="3.19",
        help='电脑管家版本号，用于计算转换按钮位置；3.19及以上使用 190，低于3.19 使用 210'
    )

    parser.add_argument(
        '--done-offset',
        dest='done_button_offset',
        type=int,
        default=None,
        help='转换按钮右侧偏移量（像素）。设置后优先生效，留空则按电脑管家版本推断'
    )

    parser.add_argument(
        '--pages', '-p',
        dest='pages',
        type=str,
        default=None,
        help='页范围，格式示例: 1-3,5,7- （与 Word 打印页范围一致）'
    )

    parser.add_argument(
        '--no-calibrate',
        dest='capture_done_offset',
        action='store_false',
        help='禁用首次页面的手动点击校准（默认启用）'
    )
    
    args = parser.parse_args()
    
    # 配置参数
    pdf_file = args.pdf_file
    pdf_name = Path(pdf_file).stem
    
    # 定义目录
    workspace_dir = Path(args.output) if args.output else Path("workspace")
    png_dir = workspace_dir / f"{pdf_name}_pngs"
    ppt_dir = workspace_dir / f"{pdf_name}_ppt"
    out_ppt_file = workspace_dir / f"{pdf_name}.pptx"
    workspace_dir.mkdir(exist_ok=True, parents=True)

    ratio = min(screen_width/16, screen_height/9)
    max_display_width = int(16 * ratio)
    max_display_height = int(9 * ratio)

    display_width = int(max_display_width * args.size_ratio)
    display_height = int(max_display_height * args.size_ratio)
    
    print("=" * 60)
    print("NotebookLM2PPT - 将 PDF 转换为可编辑 PowerPoint 演示文稿")
    print("=" * 60)
    print(f"PDF 文件: {pdf_file}")
    print(f"输出目录: {workspace_dir}")
    print(f"图像修复（去水印）: {'启用' if args.inpaint else '禁用'}")
    print(f"DPI: {args.dpi}")
    print(f"延迟时间: {args.delay} 秒")
    print(f"超时时间: {args.timeout} 秒")
    print(f"显示尺寸: {display_width}x{display_height} (比例: {args.size_ratio})")
    print(f"电脑管家版本: {args.pc_manager_version or '未指定 (使用偏移默认值)'}")
    print(f"完成按钮偏移: {args.done_button_offset if args.done_button_offset is not None else '按版本自动推断'}")
    print("=" * 60)
    print()

    # 解析页范围参数（如果有）
    def _parse_page_range(range_str):
        if not range_str:
            return None
        pages = set()
        parts = [p.strip() for p in range_str.split(',') if p.strip()]
        for part in parts:
            if '-' in part:
                start_end = part.split('-')
                if start_end[0] == '':
                    continue
                start = int(start_end[0])
                if start_end[1] == '':
                    # 开放式范围，例如 7-
                    # 表示从 start 到最后，表示为 None, 主逻辑会处理为 None（不限制）
                    # 这里将把开放式范围当作起始页及后续页：使用 a large sentinel
                    pages.update(range(start, start + 10000))
                else:
                    end = int(start_end[1])
                    if end < start:
                        continue
                    pages.update(range(start, end + 1))
            else:
                pages.add(int(part))
        return sorted(pages)

    pages_list = _parse_page_range(args.pages)

    process_pdf_to_ppt(
        pdf_path=pdf_file,
        png_dir=png_dir,
        ppt_dir=ppt_dir,
        delay_between_images=args.delay,
        inpaint=args.inpaint,
        dpi=args.dpi,
        timeout=args.timeout,
        display_height=display_height,
        display_width=display_width,
        pc_manager_version=args.pc_manager_version,
        done_button_offset=args.done_button_offset,
        capture_done_offset=args.capture_done_offset,
        pages=pages_list,
    )

    combine_ppt(ppt_dir, out_ppt_file)
    out_ppt_file = os.path.abspath(out_ppt_file)
    os.startfile(out_ppt_file)
    print(f"\n最终合并的PPT文件: {out_ppt_file}")


if __name__ == "__main__":
    main()
