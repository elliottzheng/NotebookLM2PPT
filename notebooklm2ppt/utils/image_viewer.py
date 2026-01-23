# -*- coding: utf-8 -*-
# D:/research/ICCV/vis_pixel_fake_rebuttal3/具象抽象_模型建构_Page1.png

# 使用 PIL + Tkinter 打开图片，按屏幕比例缩放并全屏显示；无边框/工具栏。

import os
import sys
import time
import ctypes
import tkinter
from PIL import Image, ImageTk


def _get_screen_resolution():
    # Windows 获取屏幕分辨率，并启用 DPI 感知避免缩放影响
    try:
        user32 = ctypes.windll.user32
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass
        return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
    except Exception:
        return 1920, 1080


def show_image_fullscreen(image_path: str, display_height: int = None, stop_event=None, ready_event=None, stop_callback=None):
    """
    显示图片在屏幕左上角

    Args:
        image_path: 图片路径
        display_height: 指定显示高度（像素），如果为None则自动适配屏幕
        stop_event: threading.Event对象，当设置时关闭窗口
        ready_event: threading.Event对象，当窗口准备好后设置此事件
        stop_callback: 可调用对象，按ESC键时会调用此回调函数来停止整个转换流程
    """
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"未找到图片: {image_path}")

    # 使用PIL打开图片
    img = Image.open(image_path).convert("RGB")

    screen_w, screen_h = _get_screen_resolution()
    img_w, img_h = img.size

    # 根据指定高度或屏幕大小计算缩放比例
    if display_height is not None:
        scale = display_height / img_h
    else:
        scale = min(screen_w / img_w, screen_h / img_h)

    new_w = max(1, int(img_w * scale))
    new_h = max(1, int(img_h * scale))

    # 使用PIL调整图片大小
    if scale != 1.0:
        # LANCZOS 是高质量的重采样滤波器
        img = img.resize((new_w, new_h), Image.LANCZOS)

    # 检查是否已有Tkinter root，如果有则使用Toplevel，否则创建新的Tk
    is_toplevel = False
    try:
        # 尝试获取默认root
        default_root = tkinter._default_root
        if default_root and default_root.winfo_exists():
            # 已存在root，使用Toplevel
            root = tkinter.Toplevel()
            is_toplevel = True
        else:
            # 没有root或root已销毁，创建新的Tk
            root = tkinter.Tk()
            is_toplevel = False
    except (AttributeError, tkinter.TclError):
        # 如果出错，创建新的Tk
        root = tkinter.Tk()
        is_toplevel = False

    # 设置无边框全屏
    root.overrideredirect(1)
    root.geometry("%dx%d+0+0" % (screen_w, screen_h))
    root.focus_set()

    # 创建画布并填充背景为黑色
    # 去除所有边框
    canvas = tkinter.Canvas(root, width=screen_w, height=screen_h, highlightthickness=0, borderwidth=0)
    canvas.pack()
    canvas.configure(background='black')

    # 将 PIL 图片转换为 Tkinter 图片对象
    # 必须保持对图片的引用，否则会被垃圾回收
    tk_image = ImageTk.PhotoImage(img)

    # 在画布左上角显示图片（不居中）
    # 图片锚点设置为左上角(nw=northwest)，坐标为(0, 0)
    imagesprite = canvas.create_image(0, 0, anchor='nw', image=tk_image)

    # 保持对图片的引用，防止被垃圾回收
    canvas.image = tk_image

    # 添加ESC键响应：按ESC停止转换并关闭窗口
    def on_escape(event):
        print("\n⚠ 检测到ESC键，正在停止转换...")
        # 触发停止回调（停止整个转换流程）
        if stop_callback is not None:
            try:
                stop_callback()
            except Exception as e:
                print(f"调用停止回调失败: {e}")
        # 触发停止事件
        if stop_event is not None:
            stop_event.set()
        # 关闭窗口
        try:
            if not is_toplevel:
                root.quit()
            root.destroy()
        except:
            pass

    # 绑定ESC键
    root.bind('<Escape>', on_escape)

    # 置顶窗口（Windows）
    try:
        root.attributes('-topmost', True)
        root.lift()
        root.focus_force()
    except Exception:
        pass

    # 强制更新窗口以确保显示
    root.update_idletasks()
    root.update()

    # 通知窗口已经准备好
    if ready_event is not None:
        def signal_ready():
            # 等待一小段时间确保窗口完全渲染
            ready_event.set()
        # 延迟500ms后通知准备就绪，确保窗口已经完全显示
        root.after(500, signal_ready)

    # 如果提供了stop_event，定期检查是否需要关闭窗口
    if stop_event is not None:
        def check_stop_event():
            if stop_event.is_set():
                try:
                    # Toplevel窗口只需要destroy，不能调用quit()
                    # quit()会退出整个Tkinter事件循环，导致主GUI也退出
                    if not is_toplevel:
                        root.quit()
                    root.destroy()
                except:
                    pass
            else:
                # 每50ms检查一次
                try:
                    root.after(50, check_stop_event)
                except:
                    pass

        # 启动检查
        try:
            root.after(50, check_stop_event)
        except:
            pass

    # 启动mainloop
    # 注意：Toplevel窗口不需要调用mainloop，因为主事件循环已在运行
    # 在Toplevel模式下，线程会保持活动直到窗口被destroy
    if not is_toplevel:
        try:
            root.mainloop()
        except:
            pass
    else:
        # Toplevel模式：不阻塞，让主事件循环处理窗口
        # 线程会在这里等待，直到窗口被destroy
        try:
            # 使用一个简单的循环来保持线程活动
            while root.winfo_exists():
                time.sleep(0.1)
        except:
            pass


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "具象抽象_模型建构_Page1.png"
    # 如果提供了第二个参数，则作为显示高度
    height = int(sys.argv[2]) if len(sys.argv) > 2 else None
    show_image_fullscreen(path, display_height=height)