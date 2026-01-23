"""自动化微软电脑管家的智能圈选功能。

步骤:
1) 发送 Ctrl+Shift+A 打开智能圈选工具
2) 从左上角拖动到右下角选择全屏
3) 点击完成按钮以保存截图为 PPT

运行此脚本时，需要确保要捕获的屏幕可见。

注意：默认使用微软电脑管家的智能圈选功能（快捷键：Ctrl+Shift+A）
"""

import re
import time
import threading
import win32api
import win32gui
import win32con
# 不要在模块级别导入 pywinauto，避免与主 GUI 冲突，导致 https://github.com/elliottzheng/NotebookLM2PPT/issues/8
# from pywinauto import mouse, keyboard
from pathlib import Path
from tkinter import messagebox
import tkinter as tk
# Get screen dimensions
screen_width = win32api.GetSystemMetrics(0)
screen_height = win32api.GetSystemMetrics(1)


def _wait_for_left_click(timeout: float = 60.0, stop_flag=None):
    """等待用户进行一次左键点击，返回点击时的屏幕坐标 (x, y)。超时返回 None。
    
    Args:
        timeout: 超时时间（秒）
        stop_flag: 停止标志函数，返回 True 时中断等待
    """
    print(f'请手动点击"完成"按钮以保存第一页（等待 {int(timeout)} 秒）...')
    start = time.time()
    prev_state = bool(win32api.GetAsyncKeyState(0x01) & 0x8000)
    while time.time() - start < timeout:
        if stop_flag and stop_flag():
            print("检测到停止请求，中断等待")
            return None
        state = bool(win32api.GetAsyncKeyState(0x01) & 0x8000)
        # 检测按下的瞬间
        if state and not prev_state:
            # 记录按下时的坐标
            x, y = win32api.GetCursorPos()
            # 等待释放
            while bool(win32api.GetAsyncKeyState(0x01) & 0x8000):
                time.sleep(0.01)
            print(f"检测到左键点击坐标: {x}, {y}")
            return x, y
        prev_state = state
        time.sleep(0.01)
    print("等待用户点击超时。")
    return None


def get_ppt_windows():
    """获取当前所有PowerPoint窗口的句柄列表"""
    ppt_windows = []
    
    def enum_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            # PowerPoint 窗口类名通常为 "PPTFrameClass"
            if "PPTFrameClass" in class_name or "PowerPoint" in window_text:
                results.append(hwnd)
        return True
    
    win32gui.EnumWindows(enum_callback, ppt_windows)
    return ppt_windows


def get_explorer_windows():
    """获取当前所有文件资源管理器窗口的句柄列表"""
    explorer_windows = []
    
    def enum_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            # 文件资源管理器窗口类名通常为 "CabinetWClass"
            if "CabinetWClass" in class_name:
                results.append((hwnd, window_text))
        return True
    
    win32gui.EnumWindows(enum_callback, explorer_windows)
    return explorer_windows


def check_new_ppt_window(initial_windows, timeout=30, check_interval=1, stop_flag=None):
    """
    检查是否出现新的PPT窗口
    
    参数:
        initial_windows: 初始的PPT窗口句柄列表
        timeout: 超时时间（秒），默认30秒
        check_interval: 检查间隔（秒），默认1秒
        stop_flag: 停止标志函数，返回 True 时中断等待
    
    返回:
        (bool, list, str): (是否找到新窗口, 新窗口句柄列表, PPT文件名)
    """
    print(f"\n开始监测新的PowerPoint窗口 (超时时间: {timeout}秒)...")
    start_time = time.time()
    detected_new_window = False
    last_loading_window = None  # 最后一个"正在打开"的窗口
    seen_windows = set(initial_windows)  # 追踪所有见过的窗口
    
    while time.time() - start_time < timeout:
        if stop_flag and stop_flag():
            print("检测到停止请求，中断PPT窗口检测")
            return False, [], None
        
        current_windows = get_ppt_windows()
        new_windows = [w for w in current_windows if w not in seen_windows]
        
        # 更新已见过的窗口列表
        seen_windows.update(new_windows)
        
        if new_windows or detected_new_window:
            if new_windows and not detected_new_window:
                elapsed = time.time() - start_time
                print(f"✓ 检测到 {len(new_windows)} 个新的PowerPoint窗口 (耗时: {elapsed:.1f}秒)")
                detected_new_window = True
            
            # 检查所有当前窗口（不只是新窗口），因为窗口标题可能会更新
            all_new_windows = [w for w in current_windows if w not in initial_windows]
            
            for hwnd in all_new_windows:
                try:
                    window_text = win32gui.GetWindowText(hwnd)
                except:
                    continue
                
                # 检查是否是临时加载状态
                is_loading = window_text and ("正在打开" in window_text or "Opening" in window_text)
                
                if is_loading:
                    if hwnd != last_loading_window:
                        last_loading_window = hwnd
                        print(f"  - 检测到窗口正在加载: {window_text}，等待完全加载...")
                    continue
                
                # 找到有效的文件名（非空且不是加载状态）
                # 排除只有"PowerPoint"而没有文件名的情况
                if window_text and window_text.strip():
                    # 如果窗口标题只是"PowerPoint"，说明文件名还没有加载，继续等待
                    if window_text.strip().lower() == "powerpoint":
                        if hwnd != last_loading_window:
                            last_loading_window = hwnd
                            print(f"  - 窗口标题尚未完全加载（仅显示'PowerPoint'），继续等待...")
                        continue
                    
                    print(f"  ✓ 窗口加载完成: {window_text}")
                    
                    # 如果是SmartCopy窗口，识别后自动关闭
                    if "smartcopy" in window_text.lower():
                        try:
                            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                            print(f"  → SmartCopy窗口已识别并关闭")
                        except Exception as e:
                            print(f"  → 关闭SmartCopy窗口失败: {e}")
                    
                    return True, all_new_windows, window_text
        
        remaining = timeout - (time.time() - start_time)
        if remaining > 0:
            if detected_new_window:
                print(f"  等待窗口标题更新... (剩余: {remaining:.0f}秒)", end='\r')
            else:
                print(f"  等待中... (剩余: {remaining:.0f}秒)", end='\r')
            time.sleep(check_interval)
    
    # 超时了，但如果检测到了"正在打开"的窗口，返回成功但文件名为None
    # 这样调用方可以尝试查找最近的文件
    if detected_new_window:
        print(f"\n⚠ 窗口标题未更新，将尝试查找最近的PPT文件")
        all_new_windows = [w for w in get_ppt_windows() if w not in initial_windows]
        return True, all_new_windows, None
    
    print(f"\n✗ 在 {timeout} 秒内未检测到新的PowerPoint窗口")
    return False, [], None


def check_and_close_download_folder(initial_explorer_windows, timeout=10, check_interval=0.5, stop_flag=None):
    """
    检查是否出现新的文件资源管理器窗口（特别是下载文件夹），如果有则关闭
    
    参数:
        initial_explorer_windows: 初始的文件资源管理器窗口列表 [(hwnd, title), ...]
        timeout: 超时时间（秒），默认10秒
        check_interval: 检查间隔（秒），默认0.5秒
        stop_flag: 停止标志函数，返回 True 时中断等待
    
    返回:
        int: 关闭的窗口数量
    """
    print(f"\n开始监测新的文件资源管理器窗口 (超时时间: {timeout}秒)...")
    start_time = time.time()
    closed_count = 0
    initial_hwnds = [hwnd for hwnd, _ in initial_explorer_windows]
    
    while time.time() - start_time < timeout:
        if stop_flag and stop_flag():
            print("检测到停止请求，中断文件资源管理器窗口检测")
            return closed_count
        
        current_windows = get_explorer_windows()
        new_windows = [(hwnd, title) for hwnd, title in current_windows if hwnd not in initial_hwnds]
        
        if new_windows:
            for hwnd, title in new_windows:
                try:
                    # 检查是否是下载文件夹（标题通常包含"下载"或"Downloads"）
                    is_download_folder = "下载" in title or "Downloads" in title
                    
                    print(f"✓ 检测到新的文件资源管理器窗口: {title}")
                    if is_download_folder:
                        print(f"  → 检测到下载文件夹，正在关闭...")
                    
                    # 关闭新窗口（无论是否是下载文件夹，都关闭新打开的资源管理器）
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    closed_count += 1
                    print(f"  → 已发送关闭指令")
                    
                    # 将已处理的窗口加入初始列表，避免重复处理
                    initial_hwnds.append(hwnd)
                    
                except Exception as e:
                    print(f"  → 关闭窗口失败: {e}")
        
        remaining = timeout - (time.time() - start_time)
        if remaining > 0:
            time.sleep(check_interval)
    
    if closed_count > 0:
        print(f"\n✓ 共关闭 {closed_count} 个文件资源管理器窗口")
    else:
        print(f"\n✓ 未检测到新的文件资源管理器窗口")
    
    return closed_count


def create_topmost_dialog():    
    # 创建一个简单的Tkinter窗口    
    root = tk.Tk()    
    root.withdraw()  # 隐藏主窗口    
    # 设置窗口为置顶    
    root.attributes("-topmost", True)   
    msg = (
            "正在进行按钮位置校准，需要您的一次交互确认，请仔细阅读此说明\n\n"
            '点击"确定"后不要动鼠标，等到出现智能圈选工具栏后，请您移动鼠标手动点击"智能复制到PPT"按钮即可完整校准。\n\n'
            '请在准备好后点击"确定"继续。转换过程中请不要最小化窗口或干扰鼠标操作。'
    )
    a = messagebox.askokcancel('提示', msg, parent=root)
    root.destroy() 



def take_fullscreen_snip(
    delay_before_hotkey: float = 1.0,
    drag_duration: float = 3,
    click_duration: float = 0.1,
    check_ppt_window: bool = True,
    ppt_check_timeout: float = 30,
    width: int = screen_width,
    height: int = screen_height,
    done_button_right_offset: int | None = None,
    stop_flag=None,
):
    """使用微软电脑管家的智能圈选功能进行全屏截图。

    Args:
        delay_before_hotkey: 发送 Ctrl+Shift+A 之前等待的秒数
        drag_duration: 拖动操作持续的秒数（模拟等待）
        click_duration: 点击完成按钮的秒数
        check_ppt_window: 是否检查新的PPT窗口，默认True
        ppt_check_timeout: PPT窗口检测超时时间（秒），默认30
        width: 截图宽度，默认为屏幕宽度
        height: 截图高度，默认为屏幕高度
        done_button_right_offset: 完成按钮的右侧偏移量（像素），用于手动覆盖
        stop_flag: 停止标志函数，返回 True 时中断等待
        pc_manager_version: 电脑管家版本号；3.19及以上自动使用 190，低于3.19 使用 210
    
    Returns:
        tuple: (bool, str, int|None) - (是否成功检测到新窗口, PPT文件名, 已保存的偏移或 None)
               如果不需要检查PPT窗口，返回 (True, None, None)
    """

    # 延迟导入 pywinauto，避免在模块加载时就导入（会与主 GUI 冲突）
    from pywinauto import mouse, keyboard

    # 记录点击前的PPT窗口和文件资源管理器窗口
    initial_ppt_windows = get_ppt_windows() if check_ppt_window else []
    initial_explorer_windows = get_explorer_windows()
    
    if check_ppt_window:
        print(f"点击前PPT窗口数量: {len(initial_ppt_windows)}")
    print(f"点击前文件资源管理器窗口数量: {len(initial_explorer_windows)}")

    if stop_flag and stop_flag():
        print("检测到停止请求，中断截图操作")
        return False, None, None

    time.sleep(delay_before_hotkey)

    if stop_flag and stop_flag():
        print("检测到停止请求，中断截图操作")
        return False, None, None


    # 简化逻辑：优先使用函数参数；如果未提供或强制重新捕获，则要求手动点击以捕获偏移并保存
    resolved_offset = None
    computed_offset = None
    if done_button_right_offset is not None:
        resolved_offset = int(done_button_right_offset)
        print(f"使用传入的完成按钮偏移: {resolved_offset}")
    else:
        print("未传入完成按钮偏移，稍后将要求手动点击以捕获并保存偏移。")
        create_topmost_dialog()

    if stop_flag and stop_flag():
        print("检测到停止请求，中断截图操作")
        return False, None, None

    keyboard.send_keys('^+a')
    time.sleep(2)

    if stop_flag and stop_flag():
        print("检测到停止请求，中断截图操作")
        return False, None, None

    # Define key points for the snip and confirmation click.
    # top_left = (5, 5)
    top_left = (0,0)
    # delta = 4  # Small offset to ensure full coverage
    delta = int(width / 512 * 4)
    bottom_right = (width+delta, height)

    center = (width // 2, height // 2)

    print(bottom_right, width)

    # Perform the drag operation
    # Move to start position
    mouse.move(coords=top_left)
    
    # Press left button
    mouse.press(button='left', coords=top_left)
    
    # Wait for the duration to simulate the drag time

    time.sleep(1)
    


    # Release left button
    mouse.release(button='left', coords=bottom_right)

    if stop_flag and stop_flag():
        print("检测到停止请求，中断截图操作")
        return False, None, None

    if resolved_offset is None:
        

        
        
        coords = _wait_for_left_click(timeout=60, stop_flag=stop_flag)
        if coords:
            click_x, click_y = coords
            computed_offset = int((bottom_right[0]) - click_x)
            print(f"已计算并保存完成按钮偏移: {computed_offset}")
            resolved_offset = computed_offset
        else:
            print("首次偏移捕获超时或未检测到点击，放弃操作。")
            return False, None, None
    else:
        # 已有偏移，执行自动点击
        done_button = (bottom_right[0] - resolved_offset, height + 35)
        if done_button[1] > screen_height:
            done_button = (done_button[0], height - 35)
        mouse.move(coords=done_button)
        time.sleep(1)
        mouse.click(button='left', coords=done_button)
    
    # 检查是否出现新的PPT窗口
    if check_ppt_window:
        success, new_windows, ppt_filename = check_new_ppt_window(initial_ppt_windows, timeout=ppt_check_timeout, stop_flag=stop_flag)
        
        check_and_close_download_folder(initial_explorer_windows, timeout=10, stop_flag=stop_flag)
        
        return success, ppt_filename, computed_offset
    
    return True, None, None

if __name__ == "__main__":
    from .image_viewer import show_image_fullscreen

    image_path = "Hackathon_Architect_Playbook_pngs/page_0001.png"

    stop_event = threading.Event()
    ready_event = threading.Event()

    def _viewer():
        # 打开全屏窗口（传入stop_event和ready_event）
        show_image_fullscreen(image_path, stop_event=stop_event, ready_event=ready_event)

    t = threading.Thread(target=_viewer, name="tkinter_viewer", daemon=True)
    t.start()

    # 等待窗口准备好
    print("等待图片窗口显示...")
    if ready_event.wait(timeout=10):
        print("✓ 图片窗口已显示")
        time.sleep(0.5)
    else:
        print("⚠ 窗口显示超时")

    try:
        take_fullscreen_snip()
    finally:
        # 通知关闭窗口并等待线程退出
        stop_event.set()
        t.join(timeout=2)
