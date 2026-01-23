import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import sys
import os
import windnd
from pathlib import Path
from .cli import process_pdf_to_ppt
from .utils.ppt_combiner import combine_ppt, create_ppt_from_images
from .utils.screenshot_automation import screen_width, screen_height
from .utils.ppt_refiner import refine_ppt
from .utils.image_inpainter import get_method_names, METHOD_ID_TO_NAME, get_method_name_from_id
from .pdf2png import pdf_to_png
import json
import ctypes
import webbrowser
from . import __version__

MINERU_URL = "https://mineru.net/"
GITHUB_URL = "https://github.com/elliottzheng/NotebookLM2PPT"


CONFIG_FILE = Path("./config.json")


BASE_WINDOWS_DPI = 85


def enable_windows_dpi_awareness(root=None):
    """Enable DPI awareness on Windows and adjust Tk scaling.

    Call this before creating UI or immediately after creating `root`.
    This helps fix display issues when Windows scaling is 200%.
    """
    if sys.platform != "win32":
        return
    # Try modern APIs first, fall back gracefully
    try:
        user32 = ctypes.windll.user32
        # Try SetProcessDpiAwarenessContext (Windows 10+)
        if hasattr(user32, 'SetProcessDpiAwarenessContext'):
            try:
                # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
                user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
            except Exception:
                pass
        else:
            # Try shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE=2)
            try:
                shcore = ctypes.windll.shcore
                shcore.SetProcessDpiAwareness(2)
            except Exception:
                # Older fallback
                try:
                    user32.SetProcessDPIAware()
                except Exception:
                    pass

        # If a root was provided, try to set tk scaling to system DPI
        if root is not None:
            try:
                # Get system DPI (fallback to 96)
                dpi = BASE_WINDOWS_DPI
                
                if hasattr(user32, 'GetDpiForSystem'):
                    dpi = user32.GetDpiForSystem()
                elif hasattr(user32, 'GetDeviceCaps'):
                    # Last resort: get DC dpi
                    hdc = user32.GetDC(0)
                    # LOGPIXELSX = 88
                    gdi32 = ctypes.windll.gdi32
                    dpi = gdi32.GetDeviceCaps(hdc, 88)
                scaling = float(dpi) / BASE_WINDOWS_DPI
                print(f"系统 DPI: {dpi}, 缩放因子: {scaling}")
                root.tk.call('tk', 'scaling', scaling)
            except Exception:
                pass
    except Exception:
        pass


class TextRedirector:
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state='normal')
        self.widget.insert(tk.END, str, (self.tag,))
        self.widget.see(tk.END)
        self.widget.configure(state='disabled')

    def flush(self):
        pass

class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"NotebookLM2PPT v{__version__} - PDF 转 PPT 工具")
        self.root.geometry("850x800")
        self.root.minsize(750, 800)
        self.center_window()
        
        self.stop_flag = False
        
        self.show_startup_dialog()
        
        self.setup_ui()
        
        # Save original stdout/stderr
        self.old_stdout = sys.stdout
        self.old_stderr = sys.stderr
        
        # Redirect stdout and stderr
        sys.stdout = TextRedirector(self.log_area, "stdout")
        sys.stderr = TextRedirector(self.log_area, "stderr")
        
        if windnd:
            windnd.hook_dropfiles(self.root, func=self.on_drop_files)
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_drop_files(self, files):
        if files:
            file_path = files[0].decode('gbk') if isinstance(files[0], bytes) else files[0]
            lower_file_path = file_path.lower()
            if lower_file_path.endswith('.pdf'):
                self.pdf_path_var.set(file_path)
                print(f"已添加文件: {file_path}")
            elif lower_file_path.endswith('.json'):
                self.mineru_json_var.set(file_path)
            else:
                messagebox.showwarning("提示", "请拖拽 PDF 文件或者 Mineru JSON 文件！")

    def on_closing(self):
        self.dump_config_to_disk()
        sys.stdout = self.old_stdout
        sys.stderr = self.old_stderr
        self.root.destroy()

    def center_window(self):
        """将窗口居中显示"""
        width = 850
        height = 800
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def center_toplevel(self, window, width, height):
        """将 Toplevel 窗口居中显示"""
        x = (window.winfo_screenwidth() // 2) - (width // 2)
        y = (window.winfo_screenheight() // 2) - (height // 2)
        window.geometry(f'{width}x{height}+{x}+{y}')

    def add_context_menu(self, widget):
        """为输入框添加右键菜单（剪切、复制、粘贴、全选）"""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="剪切", command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="复制", command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="粘贴", command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="全选", command=lambda: widget.select_range(0, tk.END))
        
        def show_menu(event):
            menu.post(event.x_root, event.y_root)
        
        widget.bind("<Button-3>", show_menu)

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)

        # File Selection
        file_frame = ttk.LabelFrame(main_frame, text="📁 文件设置（支持拖拽 PDF/ 对应的MinerU JSON 文件到窗口）", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="PDF 文件:").grid(row=0, column=0, sticky=tk.W)
        self.pdf_path_var = tk.StringVar()
        pdf_entry = ttk.Entry(file_frame, textvariable=self.pdf_path_var, width=60)
        pdf_entry.grid(row=0, column=1, padx=5, sticky="ew")
        self.add_context_menu(pdf_entry)
        ttk.Button(file_frame, text="浏览...", command=self.browse_pdf).grid(row=0, column=2)

        ttk.Label(file_frame, text="输入PDF对应的MinerU JSON (可选，进一步优化效果):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.mineru_json_var = tk.StringVar(value="")
        mineru_entry = ttk.Entry(file_frame, textvariable=self.mineru_json_var, width=60)
        mineru_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.add_context_menu(mineru_entry)
        ttk.Button(file_frame, text="浏览...", command=self.browse_json).grid(row=2, column=2, pady=5)
        ttk.Button(file_frame, text="说明", command=self.show_mineru_info).grid(row=2, column=3, pady=5, padx=5)

        ttk.Label(file_frame, text="输出目录:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_dir_var = tk.StringVar(value="workspace")
        output_entry = ttk.Entry(file_frame, textvariable=self.output_dir_var, width=60)
        output_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.add_context_menu(output_entry)
        ttk.Button(file_frame, text="浏览...", command=self.browse_output).grid(row=1, column=2, pady=5)
        ttk.Button(file_frame, text="打开", command=self.open_output_dir).grid(row=1, column=3, pady=5, padx=5)        

        # Options
        opt_frame = ttk.LabelFrame(main_frame, text="⚙️ 转换选项", padding="10")
        opt_frame.pack(fill=tk.X, pady=5)
        opt_frame.columnconfigure(1, weight=1)
        opt_frame.columnconfigure(3, weight=1)
        opt_frame.columnconfigure(5, weight=1)

        # 第一行：DPI 和 等待时间
        ttk.Label(opt_frame, text="图片清晰度 (DPI):").grid(row=0, column=0, sticky=tk.W)
        self.dpi_var = tk.IntVar(value=150)
        dpi_entry = ttk.Entry(opt_frame, textvariable=self.dpi_var, width=8)
        dpi_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.add_context_menu(dpi_entry)
        ttk.Label(opt_frame, text="（150-300）", foreground="gray").grid(row=0, column=2, sticky=tk.W, padx=5)

        ttk.Label(opt_frame, text="等待时间 (秒):").grid(row=0, column=3, sticky=tk.W, padx=(20, 0))
        self.delay_var = tk.IntVar(value=2)
        delay_entry = ttk.Entry(opt_frame, textvariable=self.delay_var, width=8)
        delay_entry.grid(row=0, column=4, sticky=tk.W, padx=5)
        self.add_context_menu(delay_entry)
        ttk.Label(opt_frame, text="（每页加载后）", foreground="gray").grid(row=0, column=5, sticky=tk.W, padx=5)

        # 第二行：超时时间 和 窗口显示比例
        ttk.Label(opt_frame, text="超时时间 (秒):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.timeout_var = tk.IntVar(value=50)
        timeout_entry = ttk.Entry(opt_frame, textvariable=self.timeout_var, width=8)
        timeout_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        self.add_context_menu(timeout_entry)
        ttk.Label(opt_frame, text="（单页最大处理）", foreground="gray").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)

        ttk.Label(opt_frame, text="窗口显示比例:").grid(row=1, column=3, sticky=tk.W, pady=5, padx=(20, 0))
        self.ratio_var = tk.DoubleVar(value=0.8)
        ratio_entry = ttk.Entry(opt_frame, textvariable=self.ratio_var, width=8)
        ratio_entry.grid(row=1, column=4, sticky=tk.W, padx=5, pady=5)
        self.add_context_menu(ratio_entry)
        ttk.Label(opt_frame, text="（建议 0.7-0.9）", foreground="gray").grid(row=1, column=5, sticky=tk.W, padx=5, pady=5)

        # 第三行：去除水印 和 修复方法
        self.inpaint_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="去除水印", variable=self.inpaint_var).grid(row=2, column=0, sticky=tk.W, pady=5)

        ttk.Label(opt_frame, text="图像修复方法:").grid(row=2, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        self.inpaint_method_var = tk.StringVar(value=get_method_names()[0])
        inpaint_method_combo = ttk.Combobox(opt_frame, textvariable=self.inpaint_method_var, width=16, state="readonly")
        inpaint_method_combo['values'] = get_method_names()
        inpaint_method_combo.grid(row=2, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Button(opt_frame, text="说明", command=self.show_inpaint_method_info).grid(row=2, column=3, pady=5, padx=5)

        # 第四行：仅图片模式 和 强制重新生成
        self.image_only_var = tk.BooleanVar(value=False)
        self.image_only_var.trace_add('write', self.on_image_only_changed)
        ttk.Checkbutton(opt_frame, text="仅图片模式（跳过智能圈选，PPT不可编辑）", variable=self.image_only_var).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)

        self.force_regenerate_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_frame, text="强制重新生成所有页面", variable=self.force_regenerate_var).grid(row=3, column=3, columnspan=3, sticky=tk.W, pady=5, padx=(20, 0))

        # 第五行：页码范围
        ttk.Label(opt_frame, text="页码范围:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.page_range_var = tk.StringVar(value="")
        page_range_entry = ttk.Entry(opt_frame, textvariable=self.page_range_var, width=20)
        page_range_entry.grid(row=4, column=1, sticky="ew", padx=5, pady=5)
        self.add_context_menu(page_range_entry)
        ttk.Label(opt_frame, text="留空=全部，示例: 1-3,5,7-9", foreground="gray").grid(row=4, column=2, sticky=tk.W, padx=5, pady=5)

        # 第六行：按钮偏移
        ttk.Label(opt_frame, text="按钮偏移 (像素):").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.done_offset_var = tk.StringVar(value="")
        done_offset_entry = ttk.Entry(opt_frame, textvariable=self.done_offset_var, width=8)
        done_offset_entry.grid(row=5, column=1, sticky=tk.W, padx=5, pady=5)
        self.add_context_menu(done_offset_entry)
        self.saved_offset_var = tk.StringVar(value="")
        ttk.Label(opt_frame, textvariable=self.saved_offset_var, foreground="blue").grid(row=5, column=2, sticky=tk.W, padx=5, pady=5)

        self.calibrate_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="校准按钮位置", variable=self.calibrate_var).grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=5)

        # 提示信息
        ttk.Label(opt_frame, text="⚠️ 核心参数：程序通过模拟鼠标点击'转换为PPT'按钮实现转换", foreground="red").grid(row=7, column=0, columnspan=6, sticky=tk.W)
        ttk.Label(opt_frame, text="   如果无法准确定位按钮位置，核心功能将无法实现！可通过勾选'校准按钮位置'进行校准", foreground="red").grid(row=8, column=0, columnspan=6, sticky=tk.W)
        ttk.Label(opt_frame, text="   提示: 程序会自动保存校准结果，下次无需重复校准", foreground="red").grid(row=9, column=0, columnspan=6, sticky=tk.W)


        # Control
        ctrl_frame = ttk.Frame(main_frame, padding="10")
        ctrl_frame.pack(fill=tk.X)

        self.start_btn = ttk.Button(ctrl_frame, text="🚀 开始转换", command=self.start_conversion)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(ctrl_frame, text="⏹️ 停止转换", command=self.stop_conversion, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        # Log Area
        log_frame = ttk.LabelFrame(main_frame, text="📋 运行日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled', height=15)
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.tag_config("stderr", foreground="red")

        self.load_config_from_disk()

    def browse_pdf(self):
        current_path = self.pdf_path_var.get().strip().strip('"')
        initial_dir = None
        
        if current_path and os.path.exists(os.path.dirname(current_path)):
            initial_dir = os.path.dirname(current_path)
        elif hasattr(self, 'last_pdf_dir') and self.last_pdf_dir and os.path.exists(self.last_pdf_dir):
            initial_dir = self.last_pdf_dir
        
        filename = filedialog.askopenfilename(
            parent=self.root,
            title="选择 PDF 文件",
            initialdir=initial_dir,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.pdf_path_var.set(filename)
            self.last_pdf_dir = os.path.dirname(filename)

    def browse_output(self):
        # 清理路径中的引号和空格
        current_dir = self.output_dir_var.get().strip().strip('"')
        initial_dir = current_dir if current_dir and os.path.exists(current_dir) else None
        
        directory = filedialog.askdirectory(
            parent=self.root,
            title="选择输出目录",
            initialdir=initial_dir
        )
        if directory:
            self.output_dir_var.set(directory)
            print(f"已设置新目录: {directory}")

    def open_output_dir(self):
        output_dir = self.output_dir_var.get().strip().strip('"')
        if not output_dir:
            messagebox.showwarning("提示", "请先设置输出目录")
            return
        
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
                print(f"已创建输出目录: {output_dir}")
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录: {str(e)}")
                return
        
        try:
            os.startfile(output_dir)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开输出目录: {str(e)}")

    def browse_json(self):
        current_path = self.mineru_json_var.get().strip().strip('"')
        initial_dir = None
        
        if current_path and os.path.exists(os.path.dirname(current_path)):
            initial_dir = os.path.dirname(current_path)
        elif hasattr(self, 'last_json_dir') and self.last_json_dir and os.path.exists(self.last_json_dir):
            initial_dir = self.last_json_dir
        
        filename = filedialog.askopenfilename(
            parent=self.root,
            title="选择 Mineru JSON 文件",
            initialdir=initial_dir,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.mineru_json_var.set(filename)
            self.last_json_dir = os.path.dirname(filename)

    def on_image_only_changed(self, *args):
        if self.image_only_var.get():
            result = messagebox.askyesno(
                "确认仅图片模式",
                "仅图片模式将：\n\n"
                "• 跳过智能圈选功能\n"
                "• 直接将去水印后的PNG图片插入PPT\n"
                "• 不生成可编辑的文本内容\n"
                "• 速度更快，但PPT内容不可编辑\n\n"
                "是否继续启用仅图片模式？",
                icon='question'
            )
            if not result:
                self.image_only_var.set(False)

    def show_inpaint_method_info(self):
        from .utils.image_inpainter import INPAINT_METHODS
        
        info_lines = ["图像修复方法说明：\n"]
        for method in INPAINT_METHODS:
            info_lines.append(f"• {method['name']}\n  {method['description']}\n")
        
        info = "".join(info_lines)
        
        top = tk.Toplevel(self.root)
        top.title("图像修复方法说明")
        self.center_toplevel(top, 600, 400)
        txt = scrolledtext.ScrolledText(top, wrap=tk.WORD, height=15)
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8,6))
        txt.insert(tk.END, info)
        txt.configure(state='disabled')
        btn_frame = ttk.Frame(top)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=8, pady=6)
        ttk.Button(btn_frame, text="关闭", command=top.destroy).pack(side=tk.LEFT, padx=6)

    def start_conversion(self):
        pdf_path = self.pdf_path_var.get().strip().strip('"')
        output_dir = self.output_dir_var.get().strip().strip('"')
        
        self.pdf_path_var.set(pdf_path)
        self.output_dir_var.set(output_dir)

        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("错误", "请先选择一个 PDF 文件")
            return

        self.stop_flag = False
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        threading.Thread(target=self.run_conversion, daemon=True).start()

    def stop_conversion(self):
        self.stop_flag = True
        print("正在停止转换...")
        self.stop_btn.config(state=tk.DISABLED)

    def dump_config_to_disk(self):
        current_method = self.inpaint_method_var.get()
        
        config_data = {
            "output_dir": self.output_dir_var.get(),
            "dpi": self.dpi_var.get(),
            "delay": self.delay_var.get(),
            "timeout": self.timeout_var.get(),
            "ratio": self.ratio_var.get(),
            "inpaint": self.inpaint_var.get(),
            "inpaint_method": current_method,
            "image_only": self.image_only_var.get(),
            "force_regenerate": self.force_regenerate_var.get(),
            "done_offset": self.done_offset_var.get(),
            "last_pdf_dir": getattr(self, 'last_pdf_dir', ''),
            "last_json_dir": getattr(self, 'last_json_dir', ''),
        }
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
            print("配置已保存到磁盘")
        except Exception as e:
            print(f"配置保存失败: {str(e)}")

    def load_config_from_disk(self):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
                self.output_dir_var.set(config_data.get("output_dir", "workspace"))
                self.dpi_var.set(config_data.get("dpi", 150))
                self.delay_var.set(config_data.get("delay", 2))
                self.timeout_var.set(config_data.get("timeout", 50))
                self.ratio_var.set(config_data.get("ratio", 0.8))
                self.inpaint_var.set(config_data.get("inpaint", True))
                
                old_method = config_data.get("inpaint_method", get_method_names()[0])
                self.inpaint_method_var.set(get_method_name_from_id(old_method))
                
                self.image_only_var.set(config_data.get("image_only", False))
                self.force_regenerate_var.set(config_data.get("force_regenerate", False))
                offset_value = config_data.get("done_offset", "")
                self.update_offset_related_gui(offset_value)
                self.last_pdf_dir = config_data.get("last_pdf_dir", '')
                self.last_json_dir = config_data.get("last_json_dir", '')
        except Exception as e:
            print(f"配置加载失败: {str(e)}")
            self.dump_config_to_disk()
            print("已创建默认配置文件")


    def update_offset_disk(self, offset_value):
        self.done_offset_var.set(str(offset_value))
        self.dump_config_to_disk()
        self.update_offset_related_gui(offset_value)

    def update_offset_related_gui(self, done_offset_value=None):
        saved = done_offset_value
        is_valid = saved is not None and saved != ""
        if is_valid:
            self.saved_offset_var.set(f"已保存: {saved}px")
            if not self.done_offset_var.get().strip():
                self.done_offset_var.set(str(saved))
        else:
            self.saved_offset_var.set("未保存: 将自动校准")
        self.calibrate_var.set(not is_valid)
        
    def show_startup_dialog(self):
        config_file = Path("./config.json")
        show_dialog = True
        
        try:
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                    show_dialog = not config_data.get("hide_startup_dialog", False)
        except Exception:
            pass
        
        if not show_dialog:
            return
        
        top = tk.Toplevel(self.root)
        top.title("欢迎使用")
        self.center_toplevel(top, 500, 300)
        top.resizable(False, False)
        
        info_frame = ttk.Frame(top, padding="20")
        info_frame.pack(fill=tk.BOTH, expand=True)
        
        info_text = (
            "本软件是免费开源的 PDF 转 PPT 工具\n\n"
            "开发者：Elliott Zheng\n\n"
            "如果您感觉本软件对您有所帮助，请在项目GitHub上给个star或是介绍给您的朋友，谢谢。\n\n"
            "本软件免费开源，如果您是以付费的方式获得本软件，那么你应该是被骗了。[○･｀Д´･○]\n\n"
            "感谢使用本工具！"
        )
        
        txt = scrolledtext.ScrolledText(info_frame, wrap=tk.WORD, height=8)
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert(tk.END, info_text)
        txt.configure(state='disabled')
        
        btn_frame = ttk.Frame(info_frame)
        btn_frame.pack(fill=tk.X, pady=(15, 0))
        
        def open_github():
            try:
                webbrowser.open_new_tab(GITHUB_URL)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开网页: {e}")
        
        def on_ok():
            top.destroy()
        
        def on_dont_show():
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
            except Exception:
                config_data = {}
            
            config_data["hide_startup_dialog"] = True
            
            try:
                with open(config_file, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, ensure_ascii=False, indent=4)
            except Exception as e:
                print(f"⚠️ 保存配置失败: {e}")
            
            top.destroy()
        
        ttk.Button(btn_frame, text="打开 GitHub 主页", command=open_github).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="不再显示", command=on_dont_show).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.RIGHT, padx=5)
        
        top.transient(self.root)
        top.grab_set()
        self.root.wait_window(top)
        
    def show_mineru_info(self):
        info = (
            "MinerU 是一个可在线使用的文档解析工具。\n\n"
            "使用步骤：\n"
            "1. 在 MinerU 网站 https://mineru.net/ 上传你的 PDF，等待解析完成。\n"
            "2. 解析完成后下载生成的 JSON 文件。\n"
            "3. 在本程序的“输入PDF对应的MinerU JSON (可选)”中选择该 JSON 文件。\n\n"
            "说明：该 JSON 包含页面结构、文本和排版等信息；本程序会利用它进一步优化输出 PPT 的图像、背景和文本，从而提升生成效果。\n\n"
            "注意：请确保 JSON 与要转换的 PDF 对应，否则优化效果可能不正确。"
        )
        top = tk.Toplevel(self.root)
        top.title("关于 MinerU")
        self.center_toplevel(top, 640, 360)
        txt = scrolledtext.ScrolledText(top, wrap=tk.WORD, height=12)
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8,6))
        txt.insert(tk.END, info)
        txt.configure(state='disabled')
        btn_frame = ttk.Frame(top)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=8, pady=6)

        def open_mineru_website():
            try:
                webbrowser.open_new_tab(MINERU_URL)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开网页: {e}")

        ttk.Button(btn_frame, text="打开 MinerU 网站", command=open_mineru_website).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_frame, text="关闭", command=top.destroy).pack(side=tk.LEFT, padx=6)
        
    def run_conversion(self):
        try:
            pdf_file = self.pdf_path_var.get()
            pdf_name = Path(pdf_file).stem
            workspace_dir = Path(self.output_dir_var.get())
            png_dir = workspace_dir / f"{pdf_name}_pngs"
            ppt_dir = workspace_dir / f"{pdf_name}_ppt"
            tmp_image_dir = workspace_dir / "tmp_images"
            
            workspace_dir.mkdir(exist_ok=True, parents=True)

            offset_raw = self.done_offset_var.get().strip()
            done_offset = None
            if offset_raw:
                try:
                    done_offset = int(offset_raw)
                except ValueError:
                    raise ValueError("完成按钮偏移需填写整数或留空")

            ratio = min(screen_width/16, screen_height/9)
            max_display_width = int(16 * ratio)
            max_display_height = int(9 * ratio)

            display_width = int(max_display_width * self.ratio_var.get())
            display_height = int(max_display_height * self.ratio_var.get())

            print(f"开始处理: {pdf_file}")

            # 解析页范围
            def parse_page_range(range_str):
                if not range_str:
                    return None
                pages = set()
                # 将中文逗号替换为英文逗号
                range_str = range_str.replace('，', ',')
                # 将各种中文破折号替换为英文连字符
                range_str = range_str.replace('—', '-').replace('–', '-').replace('－', '-')
                for part in [p.strip() for p in range_str.split(',') if p.strip()]:
                    if '-' in part:
                        start_end = part.split('-')
                        if start_end[0] == '':
                            continue
                        start = int(start_end[0])
                        if start_end[1] == '':
                            pages.update(range(start, start + 10000))
                        else:
                            end = int(start_end[1])
                            if end >= start:
                                pages.update(range(start, end + 1))
                    else:
                        pages.add(int(part))
                return sorted(pages)

            # 将页码列表转换为字符串格式
            def format_page_suffix(pages):
                if not pages:
                    return ""
                result = []
                i = 0
                while i < len(pages):
                    start = pages[i]
                    end = start
                    while i + 1 < len(pages) and pages[i + 1] == end + 1:
                        i += 1
                        end = pages[i]
                    if start == end:
                        result.append(str(start))
                    else:
                        result.append(f"{start}-{end}")
                    i += 1
                return f"_p{','.join(result)}"

            pages_list = None
            try:
                pages_list = parse_page_range(self.page_range_var.get().strip())
            except Exception as e:
                raise ValueError("页范围格式错误，请使用 1-3,5,7- 类似格式")
            
            # 根据页码范围生成文件名后缀
            page_suffix = format_page_suffix(pages_list)
            out_ppt_file = workspace_dir / f"{pdf_name}{page_suffix}.pptx"
            
            if self.image_only_var.get():
                print("=" * 60)
                print("仅图片模式：直接将PNG图片插入PPT")
                print("=" * 60)
                
                png_names = pdf_to_png(
                    pdf_path=pdf_file,
                    output_dir=png_dir,
                    dpi=self.dpi_var.get(),
                    inpaint=self.inpaint_var.get(),
                    pages=pages_list,
                    inpaint_method=self.inpaint_method_var.get(),
                    force_regenerate=self.force_regenerate_var.get()
                )
                
                if self.stop_flag:
                    print("\n⏹️ 转换已被用户停止")
                    messagebox.showinfo("转换已停止", "转换已被用户停止")
                    return
                
                png_names = create_ppt_from_images(png_dir, out_ppt_file, png_names=png_names)
                extra_message = "（仅图片模式）"
            else:
                png_names = process_pdf_to_ppt(
                    pdf_path=pdf_file,
                    png_dir=png_dir,
                    ppt_dir=ppt_dir,
                    delay_between_images=self.delay_var.get(),
                    inpaint=self.inpaint_var.get(),
                    dpi=self.dpi_var.get(),
                    timeout=self.timeout_var.get(),
                    display_height=display_height,
                    display_width=display_width,
                    done_button_offset=done_offset,
                    capture_done_offset=self.calibrate_var.get(),
                    pages=pages_list,
                    update_offset_callback=self.update_offset_disk,
                    stop_flag=lambda: self.stop_flag,
                    force_regenerate=self.force_regenerate_var.get(),
                    inpaint_method=self.inpaint_method_var.get()
                )

                if self.stop_flag:
                    print("\n⏹️ 转换已被用户停止")
                    messagebox.showinfo("转换已停止", "转换已被用户停止")
                    return

                png_names = combine_ppt(ppt_dir, out_ppt_file, png_names=png_names)
                extra_message = ""
            if not self.image_only_var.get():
                # 如果用户提供了 mineru JSON，则进行 refine_ppt 处理
                mineru_json = self.mineru_json_var.get().strip().strip('"')
                if mineru_json:
                    if not os.path.exists(mineru_json):
                        print(f"⚠️ 提供的 MinerU JSON 文件不存在，跳过 PPT 优化: {mineru_json}")
                    else:
                        refined_out = workspace_dir / f"{pdf_name}{page_suffix}_优化.pptx"
                        print(f"开始利用MinerU信息优化 PPT: {mineru_json}")
                        refine_ppt(str(tmp_image_dir), mineru_json, str(out_ppt_file), str(png_dir), png_names, str(refined_out))
                        
                        print("✅ refine_ppt 完成")
                        extra_message = "优化前的PPT已保存在同一目录下"
                        out_ppt_file = os.path.abspath(refined_out)
                else:
                    extra_message = ""
            else:
                extra_message = "（仅图片模式）"
            out_ppt_file = os.path.abspath(out_ppt_file)
            print(f"\n✅ 转换完成！")
            print(f"📄 输出文件: {out_ppt_file}")
            os.startfile(out_ppt_file)
            messagebox.showinfo("转换成功", f"PDF 已成功转换为 PPT！\n\n文件位置:\n{out_ppt_file}"+extra_message)
        except Exception as e:
            print(f"\n❌ 转换失败: {str(e)}")
            messagebox.showerror("转换失败", f"处理过程中出现错误:\n{str(e)}")
        finally:
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)

def launch_gui():
    # Enable Windows DPI awareness before creating the Tk root where possible
    try:
        enable_windows_dpi_awareness(None)
    except Exception:
        pass

    root = tk.Tk()
    # After root exists, apply scaling using the helper (this will call tk scaling)
    try:
        enable_windows_dpi_awareness(root)
    except Exception:
        pass

    app = AppGUI(root)
    root.mainloop()

if __name__ == "__main__":
    launch_gui()
