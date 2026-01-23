"""命令行界面：将 PDF 转换为可编辑 PowerPoint 演示文稿"""

import os
import time
import threading
import shutil
import argparse
import sys
from pathlib import Path
from .pdf2png import pdf_to_png
from .utils.image_viewer import show_image_fullscreen
from .utils.screenshot_automation import take_fullscreen_snip, screen_height, screen_width


def process_pdf_to_ppt(pdf_path, png_dir, ppt_dir, delay_between_images=2, inpaint=True, dpi=150, timeout=50, display_height=None, 
                    display_width=None, done_button_offset=None, capture_done_offset: bool = True, pages=None, update_offset_callback=None, stop_flag=None, force_regenerate=False):
    """
    将 PDF 转换为 PNG 图片，然后对每张图片进行截图处理
    
    Args:
        pdf_path: PDF 文件路径
        png_dir: PNG 输出目录
        ppt_dir: PPT 输出目录
        delay_between_images: 每张图片之间的延迟时间（秒）
        inpaint: 是否进行图像修复
        dpi: 图片清晰度
        timeout: 超时时间（秒）
        display_height: 显示窗口高度
        display_width: 显示窗口宽度
        done_button_offset: 完成按钮右侧偏移量
        capture_done_offset: 是否捕获完成按钮偏移
        pages: 要处理的页码范围
        update_offset_callback: 偏移更新回调函数
        stop_flag: 停止标志（用于中断转换）
        force_regenerate: 是否强制重新生成所有 PPT（默认 False，复用已存在的 PPT）
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

    # 创建一个本地停止标志，用于响应ESC键
    esc_stop_requested = [False]  # 使用列表以便在嵌套函数中修改

    # 创建组合停止标志函数，同时检查外层stop_flag和ESC键停止请求
    def combined_stop_flag():
        return (stop_flag and stop_flag()) or esc_stop_requested[0]

    # 3. 对每张图片进行截图处理
    for idx, png_file in enumerate(png_files, 1):
        if combined_stop_flag():
            print("\n⏹️ 用户请求停止转换")
            break

        # 检查是否按了ESC键
        if esc_stop_requested[0]:
            print("\n⏹️ 用户按ESC键停止转换")
            break
        
        print(f"\n[{idx}/{len(png_files)}] 处理图片: {png_file.name}")
        
        target_filename = png_file.stem + ".pptx"
        target_path = ppt_dir / target_filename
        
        if not force_regenerate and target_path.exists():
            print(f"  ✓ PPT文件已存在，跳过转换: {target_path}")
            continue
        
        stop_event = threading.Event()
        ready_event = threading.Event()

        # 创建一个回调函数，用于在按ESC键时停止整个转换流程
        def on_stop_requested():
            """当用户按ESC键时，设置停止标志"""
            print("用户请求停止转换（按ESC键）")
            esc_stop_requested[0] = True

        def _viewer():
            """在线程中显示图片"""
            # 传入stop_event、ready_event和stop_callback
            show_image_fullscreen(str(png_file), display_height=display_height,
                                 stop_event=stop_event, ready_event=ready_event,
                                 stop_callback=on_stop_requested)

        # 启动图片显示线程
        viewer_thread = threading.Thread(
            target=_viewer,
            name=f"tkinter_viewer_{idx}",
            daemon=True
        )
        viewer_thread.start()

        # 等待窗口准备好（最多等待10秒）
        print("等待图片窗口显示...")
        window_ready = ready_event.wait(timeout=10)
        if not window_ready:
            print("⚠ 窗口显示超时，继续执行...")
        else:
            print("✓ 图片窗口已显示")
            # 额外等待一小段时间确保窗口稳定
            time.sleep(0.5)
        
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
                stop_flag=combined_stop_flag,
            )
            if combined_stop_flag():
                print("\n⏹️ 用户请求停止转换")
                break
            if esc_stop_requested[0]:
                print("\n⏹️ 用户按ESC键停止转换")
                break
            if success and computed_offset is not None:
                print(f"捕获到的完成按钮偏移: {computed_offset}")
                done_button_offset = computed_offset  # 更新为最新捕获的偏移
                if update_offset_callback:
                    update_offset_callback(computed_offset)


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
                # 延迟导入 pywinauto，避免在模块加载时就导入（会与主 GUI 冲突）
                from pywinauto import mouse
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

    # 删除CLI
    print("命令行模式已被弃用，请使用 GUI 界面。")
    

if __name__ == "__main__":
    main()
