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
from .i18n import get_text, SUPPORTED_LANGUAGES, set_language
from .utils.process_checker import is_process_running, PROCESS_NAME

MINERU_URL = "https://mineru.net/"
GITHUB_URL = "https://github.com/elliottzheng/NotebookLM2PPT"
PC_MANAGER_URL = "https://pcmanager.microsoft.com/"


CONFIG_FILE = Path("./config.json")


BASE_WINDOWS_DPI = 85


def icon_path():
    """获取资源绝对路径，适用于PyInstaller打包后"""
    try:
        return os.path.join(sys._MEIPASS, "favicon.ico")
    except Exception:
        return os.path.abspath("./docs/public/favicon.ico")
    
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
        self.lang = "zh_cn"  # Default
        self.top_left = (10, 10)
        self.load_config_from_disk()
        set_language(self.lang)
        
        self.root.title(get_text("root_title", version=__version__))
        self.root.geometry("850x920")
        self.root.minsize(750, 850)
        self.center_window()
        
        self.stop_flag = False
        
        self.show_startup_dialog()
        
        self.setup_ui()
        
        # Load config again to populate UI variables
        self.load_config_from_disk()
        
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
            decoded_files = []
            for f in files:
                decoded_files.append(f.decode('gbk') if isinstance(f, bytes) else f)
            pdfs = [f for f in decoded_files if f.lower().endswith('.pdf')]
            jsons = [f for f in decoded_files if f.lower().endswith('.json')]
            if len(pdfs) <= 1 and len(decoded_files) == 1:
                file_path = decoded_files[0]
                lower_file_path = file_path.lower()
                if lower_file_path.endswith('.pdf'):
                    self.pdf_path_var.set(file_path)
                    print(get_text("file_added_msg", file=file_path))
                elif lower_file_path.endswith('.json'):
                    self.mineru_json_var.set(file_path)
                else:
                    messagebox.showwarning(get_text("info_btn"), get_text("drag_drop_warning"))
                return
            if pdfs:
                json_map = {Path(j).stem: j for j in jsons}
                for p in pdfs:
                    matched_json = json_map.get(Path(p).stem, "")
                    if not matched_json and len(jsons) == 1:
                        matched_json = jsons[0]
                    self.add_task(p, matched_json or None)
            else:
                messagebox.showwarning(get_text("info_btn"), get_text("drag_drop_warning"))

    def on_closing(self):
        self.dump_config_to_disk()
        sys.stdout = self.old_stdout
        sys.stderr = self.old_stderr
        self.root.destroy()

    def center_window(self):
        """将窗口居中显示"""
        width = 850
        height = 920
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def create_toplevel(self, title, width, height):
        """创建带图标设置的 Toplevel 窗口"""
        top = tk.Toplevel(self.root)
        top.title(title)
        top.iconbitmap(icon_path())
        self.center_toplevel(top, width, height)
        return top

    def center_toplevel(self, window, width, height):
        """将 Toplevel 窗口居中显示"""
        x = (window.winfo_screenwidth() // 2) - (width // 2)
        y = (window.winfo_screenheight() // 2) - (height // 2)
        window.geometry(f'{width}x{height}+{x}+{y}')

    def change_language(self, new_lang):
        """切换语言并重启 UI"""
        if self.lang == new_lang:
            return
        
        self.lang = new_lang
        set_language(self.lang)
        self.dump_config_to_disk()
        
        # 更新主窗口标题
        self.root.title(get_text("root_title", version=__version__))
        
        # 刷新主 UI
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Frame) or isinstance(widget, tk.Frame):
                widget.destroy()
        
        self.setup_ui()
        # 重新重定向 stdout/stderr 到新的 log_area
        sys.stdout = TextRedirector(self.log_area, "stdout")
        sys.stderr = TextRedirector(self.log_area, "stderr")

    def add_context_menu(self, widget):
        """为输入框添加右键菜单（剪切、复制、粘贴、全选）"""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label=get_text("cut"), command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label=get_text("copy"), command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label=get_text("paste"), command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label=get_text("select_all"), command=lambda: widget.select_range(0, tk.END))
        
        def show_menu(event):
            menu.post(event.x_root, event.y_root)
        
        widget.bind("<Button-3>", show_menu)

    def get_translated_method_names(self):
        """获取翻译后的方法名列表"""
        from .utils.image_inpainter import INPAINT_METHODS
        return [get_text(f"method_{m['id']}_name") for m in INPAINT_METHODS]

    def get_method_id_from_translated_name(self, translated_name):
        """根据翻译后的方法名获取 ID"""
        from .utils.image_inpainter import INPAINT_METHODS
        for m in INPAINT_METHODS:
            if get_text(f"method_{m['id']}_name") == translated_name:
                return m['id']
        return "background_smooth"

    def get_translated_name_from_id(self, method_id):
        """根据 ID 获取翻译后的方法名"""
        return get_text(f"method_{method_id}_name")

    def on_language_combo_change(self, event):
        """处理语言下拉框选择变更"""
        selected_name = self.lang_combo_var.get()
        for code in SUPPORTED_LANGUAGES.keys():
            if get_text(f"lang_{code}") == selected_name:
                self.change_language(code)
                return

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)

        self.task_queue = getattr(self, 'task_queue', [])
        self.task_id_counter = getattr(self, 'task_id_counter', 1)
        self.queue_thread = None
        self.queue_stop_flag = False
        self.is_queue_running = False

        # 语言切换栏
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        
        lang_frame = ttk.Frame(header_frame)
        lang_frame.pack(side=tk.RIGHT)
        
        ttk.Label(lang_frame, text=get_text("language_menu") + ": ").pack(side=tk.LEFT)
        
        lang_display_names = [get_text(f"lang_{code}") for code in SUPPORTED_LANGUAGES.keys()]
        current_lang_display = get_text(f"lang_{self.lang}")
        
        self.lang_combo_var = tk.StringVar(value=current_lang_display)
        lang_combo = ttk.Combobox(lang_frame, textvariable=self.lang_combo_var, values=lang_display_names, state="readonly", width=10)
        lang_combo.pack(side=tk.LEFT)
        lang_combo.bind("<<ComboboxSelected>>", self.on_language_combo_change)

        # File Selection
        file_frame = ttk.LabelFrame(main_frame, text=get_text("file_settings_label"), padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text=get_text("pdf_file_label")).grid(row=0, column=0, sticky=tk.W)
        self.pdf_path_var = getattr(self, 'pdf_path_var', tk.StringVar())
        pdf_entry = ttk.Entry(file_frame, textvariable=self.pdf_path_var, width=60)
        pdf_entry.grid(row=0, column=1, padx=5, sticky="ew")
        self.add_context_menu(pdf_entry)
        ttk.Button(file_frame, text=get_text("browse_btn"), command=self.browse_pdf).grid(row=0, column=2)

        ttk.Label(file_frame, text=get_text("mineru_json_label")).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.mineru_json_var = getattr(self, 'mineru_json_var', tk.StringVar(value=""))
        mineru_entry = ttk.Entry(file_frame, textvariable=self.mineru_json_var, width=60)
        mineru_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.add_context_menu(mineru_entry)
        ttk.Button(file_frame, text=get_text("browse_btn"), command=self.browse_json).grid(row=2, column=2, pady=5)
        ttk.Button(file_frame, text=get_text("info_btn"), command=self.show_mineru_info).grid(row=2, column=3, pady=5, padx=5)

        self.unify_font_var = getattr(self, 'unify_font_var', tk.BooleanVar(value=False))
        self.font_name_var = getattr(self, 'font_name_var', tk.StringVar(value="Calibri"))
        
        self.unify_font_frame = ttk.Frame(file_frame)
        self.unify_font_frame.grid(row=3, column=1, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        self.unify_font_check = ttk.Checkbutton(self.unify_font_frame, text=get_text("unify_font_label"), variable=self.unify_font_var)
        self.unify_font_check.pack(side=tk.LEFT)
        
        self.font_label = ttk.Label(self.unify_font_frame, text=get_text("font_name_label"))
        self.font_label.pack(side=tk.LEFT, padx=(10, 2))
        
        self.font_entry = ttk.Entry(self.unify_font_frame, textvariable=self.font_name_var, width=15)
        self.font_entry.pack(side=tk.LEFT)
        self.add_context_menu(self.font_entry)

        if not hasattr(self, '_mineru_json_trace'):
            self._mineru_json_trace = self.mineru_json_var.trace_add('write', lambda *args: self.update_unify_font_visibility())
        self.update_unify_font_visibility()

        ttk.Label(file_frame, text=get_text("output_dir_label")).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_dir_var = getattr(self, 'output_dir_var', tk.StringVar(value="workspace"))
        output_entry = ttk.Entry(file_frame, textvariable=self.output_dir_var, width=60)
        output_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.add_context_menu(output_entry)
        ttk.Button(file_frame, text=get_text("browse_btn"), command=self.browse_output).grid(row=1, column=2, pady=5)
        ttk.Button(file_frame, text=get_text("open_btn"), command=self.open_output_dir).grid(row=1, column=3, pady=5, padx=5)        

        # Options
        opt_frame = ttk.LabelFrame(main_frame, text=get_text("options_label"), padding="10")
        opt_frame.pack(fill=tk.X, pady=5)
        opt_frame.columnconfigure(1, weight=1)
        opt_frame.columnconfigure(3, weight=1)
        opt_frame.columnconfigure(5, weight=1)

        # 第一行：DPI 和 等待时间
        ttk.Label(opt_frame, text=get_text("dpi_label")).grid(row=0, column=0, sticky=tk.W)
        self.dpi_var = getattr(self, 'dpi_var', tk.IntVar(value=150))
        dpi_entry = ttk.Entry(opt_frame, textvariable=self.dpi_var, width=8)
        dpi_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.add_context_menu(dpi_entry)
        ttk.Label(opt_frame, text=get_text("dpi_hint"), foreground="gray").grid(row=0, column=2, sticky=tk.W, padx=5)

        ttk.Label(opt_frame, text=get_text("delay_label")).grid(row=0, column=3, sticky=tk.W, padx=(20, 0))
        self.delay_var = getattr(self, 'delay_var', tk.IntVar(value=0))
        delay_entry = ttk.Entry(opt_frame, textvariable=self.delay_var, width=8)
        delay_entry.grid(row=0, column=4, sticky=tk.W, padx=5)
        self.add_context_menu(delay_entry)
        ttk.Label(opt_frame, text=get_text("delay_hint"), foreground="gray").grid(row=0, column=5, sticky=tk.W, padx=5)

        # 第二行：超时时间 和 窗口显示比例
        ttk.Label(opt_frame, text=get_text("timeout_label")).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.timeout_var = getattr(self, 'timeout_var', tk.IntVar(value=50))
        timeout_entry = ttk.Entry(opt_frame, textvariable=self.timeout_var, width=8)
        timeout_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        self.add_context_menu(timeout_entry)
        ttk.Label(opt_frame, text=get_text("timeout_hint"), foreground="gray").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)

        ttk.Label(opt_frame, text=get_text("ratio_label")).grid(row=1, column=3, sticky=tk.W, pady=5, padx=(20, 0))
        self.ratio_var = getattr(self, 'ratio_var', tk.DoubleVar(value=0.8))
        ratio_entry = ttk.Entry(opt_frame, textvariable=self.ratio_var, width=8)
        ratio_entry.grid(row=1, column=4, sticky=tk.W, padx=5, pady=5)
        self.add_context_menu(ratio_entry)
        ttk.Label(opt_frame, text=get_text("ratio_hint"), foreground="gray").grid(row=1, column=5, sticky=tk.W, padx=5, pady=5)

        # 第三行：去除水印 和 修复方法
        self.inpaint_var = getattr(self, 'inpaint_var', tk.BooleanVar(value=True))
        ttk.Checkbutton(opt_frame, text=get_text("inpaint_label"), variable=self.inpaint_var).grid(row=2, column=0, sticky=tk.W, pady=5)

        ttk.Label(opt_frame, text=get_text("inpaint_method_label")).grid(row=2, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        self.inpaint_method_var = getattr(self, 'inpaint_method_var', tk.StringVar(value=self.get_translated_method_names()[0]))
        inpaint_method_combo = ttk.Combobox(opt_frame, textvariable=self.inpaint_method_var, width=16, state="readonly")
        inpaint_method_combo['values'] = self.get_translated_method_names()
        inpaint_method_combo.grid(row=2, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Button(opt_frame, text=get_text("info_btn"), command=self.show_inpaint_method_info).grid(row=2, column=3, pady=5, padx=5)

        # 第四行：仅图片模式 和 强制重新生成
        self.image_only_var = getattr(self, 'image_only_var', tk.BooleanVar(value=False))
        if not hasattr(self, '_image_only_trace'):
            self._image_only_trace = self.image_only_var.trace_add('write', self.on_image_only_changed)
        ttk.Checkbutton(opt_frame, text=get_text("image_only_label"), variable=self.image_only_var).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)

        self.force_regenerate_var = getattr(self, 'force_regenerate_var', tk.BooleanVar(value=False))
        ttk.Checkbutton(opt_frame, text=get_text("force_regenerate_label"), variable=self.force_regenerate_var).grid(row=3, column=3, columnspan=3, sticky=tk.W, pady=5, padx=(20, 0))

        # 第五行：页码范围
        ttk.Label(opt_frame, text=get_text("page_range_label")).grid(row=4, column=0, sticky=tk.W, pady=5)
        self.page_range_var = getattr(self, 'page_range_var', tk.StringVar(value=""))
        page_range_entry = ttk.Entry(opt_frame, textvariable=self.page_range_var, width=20)
        page_range_entry.grid(row=4, column=1, sticky="ew", padx=5, pady=5)
        self.add_context_menu(page_range_entry)
        ttk.Label(opt_frame, text=get_text("page_range_hint"), foreground="gray").grid(row=4, column=2, sticky=tk.W, padx=5, pady=5)

        # 第六行：按钮偏移 和 自动校准
        ttk.Label(opt_frame, text=get_text("button_offset_label")).grid(row=5, column=0, sticky=tk.W, pady=5)
        self.done_offset_var = getattr(self, 'done_offset_var', tk.StringVar(value=""))
        done_offset_entry = ttk.Entry(opt_frame, textvariable=self.done_offset_var, width=8)
        done_offset_entry.grid(row=5, column=1, sticky=tk.W, padx=5, pady=5)
        self.add_context_menu(done_offset_entry)
        self.saved_offset_var = getattr(self, 'saved_offset_var', tk.StringVar(value=""))
        ttk.Label(opt_frame, textvariable=self.saved_offset_var, foreground="blue").grid(row=5, column=2, sticky=tk.W, padx=5, pady=5)

        self.calibrate_var = getattr(self, 'calibrate_var', tk.BooleanVar(value=True))
        ttk.Checkbutton(opt_frame, text=get_text("calibrate_label"), variable=self.calibrate_var).grid(row=5, column=3, columnspan=3, sticky=tk.W, pady=5, padx=(20, 0))

        # 提示信息
        ttk.Label(opt_frame, text=get_text("core_param_warning"), foreground="red").grid(row=6, column=0, columnspan=6, sticky=tk.W)
        ttk.Label(opt_frame, text=get_text("core_param_warning2"), foreground="red").grid(row=7, column=0, columnspan=6, sticky=tk.W)
        ttk.Label(opt_frame, text=get_text("core_param_warning3"), foreground="red").grid(row=8, column=0, columnspan=6, sticky=tk.W)


        # Control
        ctrl_frame = ttk.Frame(main_frame, padding="10")
        ctrl_frame.pack(fill=tk.X)

        self.start_btn = ttk.Button(ctrl_frame, text=get_text("start_btn"), command=self.start_queue)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(ctrl_frame, text=get_text("stop_btn"), command=self.stop_queue, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        queue_frame = ttk.LabelFrame(main_frame, text=get_text("queue_label"), padding="10")
        queue_frame.pack(fill=tk.BOTH, expand=False, pady=5)
        queue_frame.columnconfigure(0, weight=1)

        columns = ("id", "pdf", "json", "status", "output")
        self.queue_tree = ttk.Treeview(queue_frame, columns=columns, show="headings", height=5)
        self.queue_tree.heading("id", text=get_text("queue_col_id"))
        self.queue_tree.heading("pdf", text=get_text("queue_col_pdf"))
        self.queue_tree.heading("json", text=get_text("queue_col_json"))
        self.queue_tree.heading("status", text=get_text("queue_col_status"))
        self.queue_tree.heading("output", text=get_text("queue_col_output"))
        self.queue_tree.column("id", width=60, anchor="center")
        self.queue_tree.column("pdf", width=220)
        self.queue_tree.column("json", width=200)
        self.queue_tree.column("status", width=100, anchor="center")
        self.queue_tree.column("output", width=220)
        self.queue_tree.grid(row=0, column=0, columnspan=6, sticky="nsew", pady=(0, 8))
        self.queue_tree.bind("<Double-1>", self.on_task_double_click)

        for task in self.task_queue:
            self.queue_tree.insert("", tk.END, iid=str(task["id"]), values=(task["id"], task["pdf"], task["json"], task["status"], task["output"]))

        queue_btns = ttk.Frame(queue_frame)
        queue_btns.grid(row=1, column=0, columnspan=6, sticky="ew")
        ttk.Button(queue_btns, text=get_text("queue_add_task"), command=self.add_task_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(queue_btns, text=get_text("queue_add_multi_pdf"), command=self.add_tasks_multi_pdfs).pack(side=tk.LEFT, padx=5)
        ttk.Button(queue_btns, text=get_text("queue_remove_selected"), command=self.remove_selected_task).pack(side=tk.LEFT, padx=5)
        ttk.Button(queue_btns, text=get_text("queue_clear"), command=self.clear_tasks).pack(side=tk.LEFT, padx=5)

        # Log Area
        log_frame = ttk.LabelFrame(main_frame, text=get_text("log_area_label"), padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled', height=12)
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.tag_config("stderr", foreground="red")

    def browse_pdf(self):
        current_path = self.pdf_path_var.get().strip().strip('"')
        initial_dir = None
        
        if current_path and os.path.exists(os.path.dirname(current_path)):
            initial_dir = os.path.dirname(current_path)
        elif hasattr(self, 'last_pdf_dir') and self.last_pdf_dir and os.path.exists(self.last_pdf_dir):
            initial_dir = self.last_pdf_dir
        
        filename = filedialog.askopenfilename(
            parent=self.root,
            title=get_text("select_pdf_title"),
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
            title=get_text("select_output_title"),
            initialdir=initial_dir
        )
        if directory:
            self.output_dir_var.set(directory)
            print(get_text("set_new_dir_msg", directory=directory))

    def open_output_dir(self):
        output_dir = self.output_dir_var.get().strip().strip('"')
        if not output_dir:
            messagebox.showwarning(get_text("info_btn"), get_text("set_output_dir_warning"))
            return
        
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
                print(get_text("create_output_dir_msg", output_dir=output_dir))
            except Exception as e:
                messagebox.showerror(get_text("error_btn"), get_text("create_output_dir_error", error=str(e)))
                return
        
        try:
            os.startfile(output_dir)
        except Exception as e:
            messagebox.showerror(get_text("error_btn"), get_text("open_output_dir_error", error=str(e)))

    def browse_json(self):
        current_path = self.mineru_json_var.get().strip().strip('"')
        initial_dir = None
        
        if current_path and os.path.exists(os.path.dirname(current_path)):
            initial_dir = os.path.dirname(current_path)
        elif hasattr(self, 'last_json_dir') and self.last_json_dir and os.path.exists(self.last_json_dir):
            initial_dir = self.last_json_dir
        
        filename = filedialog.askopenfilename(
            parent=self.root,
            title=get_text("select_json_title"),
            initialdir=initial_dir,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.mineru_json_var.set(filename)
            self.last_json_dir = os.path.dirname(filename)

    def on_image_only_changed(self, *args):
        if self.image_only_var.get():
            result = messagebox.askyesno(
                get_text("image_only_confirm_title"),
                get_text("image_only_confirm_msg"),
                icon='question'
            )
            if not result:
                self.image_only_var.set(False)

    def update_unify_font_visibility(self):
        """根据是否选择 MinerU JSON 动态显示统一字体选项"""
        if self.mineru_json_var.get().strip():
            self.unify_font_frame.grid(row=3, column=1, columnspan=3, sticky=tk.W, pady=(0, 5))
        else:
            self.unify_font_frame.grid_forget()

    def show_inpaint_method_info(self):
        from .utils.image_inpainter import INPAINT_METHODS
        
        info_lines = [get_text("inpaint_method_info_prefix")]
        for method in INPAINT_METHODS:
            name = get_text(f"method_{method['id']}_name")
            desc = get_text(f"method_{method['id']}_desc")
            info_lines.append(f"• {name}\n  {desc}\n")
        
        info = "".join(info_lines)
        
        top = self.create_toplevel(get_text("inpaint_method_info_title"), 600, 400)
        txt = scrolledtext.ScrolledText(top, wrap=tk.WORD, height=15)
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8,6))
        txt.insert(tk.END, info)
        txt.configure(state='disabled')
        btn_frame = ttk.Frame(top)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=8, pady=6)
        ttk.Button(btn_frame, text=get_text("close_btn"), command=top.destroy).pack(side=tk.LEFT, padx=6)

    def ensure_pc_manager_running(self):
        if sys.platform != "win32":
            return True
        try:
            running = is_process_running(PROCESS_NAME)
        except Exception as e:
            messagebox.showerror(
                get_text("error_btn"),
                f"检测电脑管家运行状态失败: {e}",
            )
            return False
        if not running:
            full_msg = (
                get_text("pc_manager_not_running_msg")
                + "\n\n"
                + get_text("pc_manager_open_website_confirm")
            )
            open_site = messagebox.askyesno(
                get_text("error_btn"),
                full_msg,
            )
            if open_site:
                try:
                    webbrowser.open_new_tab(PC_MANAGER_URL)
                except Exception as e:
                    messagebox.showerror(
                        get_text("error_btn"),
                        get_text("open_pc_manager_website_error", error=str(e)),
                    )
            return False
        return True

    def start_conversion(self):
        pdf_path = self.pdf_path_var.get().strip().strip('"')
        output_dir = self.output_dir_var.get().strip().strip('"')
        
        self.pdf_path_var.set(pdf_path)
        self.output_dir_var.set(output_dir)

        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror(get_text("error_btn"), get_text("select_pdf_error"))
            return

        if not self.image_only_var.get():
            if not self.ensure_pc_manager_running():
                return

        self.stop_flag = False
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        threading.Thread(target=self.run_conversion, daemon=True).start()

    def stop_conversion(self):
        self.stop_flag = True
        print(get_text("stopping_msg"))
        self.stop_btn.config(state=tk.DISABLED)

    def dump_config_to_disk(self):
        current_method_name = self.inpaint_method_var.get()
        current_method_id = self.get_method_id_from_translated_name(current_method_name)
        
        config_data = {
            "language": self.lang,
            "output_dir": self.output_dir_var.get(),
            "dpi": self.dpi_var.get(),
            "delay": self.delay_var.get(),
            "timeout": self.timeout_var.get(),
            "ratio": self.ratio_var.get(),
            "inpaint": self.inpaint_var.get(),
            "inpaint_method": current_method_id,
            "image_only": self.image_only_var.get(),
            "force_regenerate": self.force_regenerate_var.get(),
            "unify_font": self.unify_font_var.get(),
            "done_offset": self.done_offset_var.get(),
            "last_pdf_dir": getattr(self, 'last_pdf_dir', ''),
            "last_json_dir": getattr(self, 'last_json_dir', ''),
        }
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
            print(get_text("config_saved"))
        except Exception as e:
            print(get_text("config_save_fail", error=str(e)))

    def load_config_from_disk(self):
        try:
            if not CONFIG_FILE.exists():
                return
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
                self.lang = config_data.get("language", "zh_cn")
                if hasattr(self, 'output_dir_var'):
                    self.output_dir_var.set(config_data.get("output_dir", "workspace"))
                    self.dpi_var.set(config_data.get("dpi", 150))
                    self.delay_var.set(config_data.get("delay", 0))
                    self.timeout_var.set(config_data.get("timeout", 50))
                    self.ratio_var.set(config_data.get("ratio", 0.8))
                    self.inpaint_var.set(config_data.get("inpaint", True))
                    
                    method_id = config_data.get("inpaint_method", "background_smooth")
                    self.inpaint_method_var.set(self.get_translated_name_from_id(method_id))
                    
                    self.image_only_var.set(config_data.get("image_only", False))
                    self.force_regenerate_var.set(config_data.get("force_regenerate", False))
                    self.unify_font_var.set(config_data.get("unify_font", False))
                    offset_value = config_data.get("done_offset", "")
                    self.update_offset_related_gui(offset_value)
                    self.last_pdf_dir = config_data.get("last_pdf_dir", '')
                    self.last_json_dir = config_data.get("last_json_dir", '')
        except Exception as e:
            print(get_text("config_load_fail", error=str(e)))
            self.dump_config_to_disk()
            print(get_text("default_config_created"))


    def update_offset_disk(self, offset_value):
        self.done_offset_var.set(str(offset_value))
        self.dump_config_to_disk()
        self.update_offset_related_gui(offset_value)

    def update_offset_related_gui(self, done_offset_value=None):
        saved = done_offset_value
        is_valid = saved is not None and saved != ""
        if is_valid:
            self.saved_offset_var.set(get_text("saved_offset", offset=saved))
            if not self.done_offset_var.get().strip():
                self.done_offset_var.set(str(saved))
        else:
            self.saved_offset_var.set(get_text("unsaved_offset"))
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
        
        top = self.create_toplevel(get_text("startup_dialog_title"), 500, 300)
        top.resizable(False, False)
        
        info_frame = ttk.Frame(top, padding="20")
        info_frame.pack(fill=tk.BOTH, expand=True)
        
        info_text = get_text("startup_info")
        
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
                messagebox.showerror(get_text("error_btn"), f"无法打开网页: {e}")
        
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
                print(get_text("config_save_fail", error=str(e)))
            
            top.destroy()
        
        ttk.Button(btn_frame, text=get_text("open_github_btn"), command=open_github).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text=get_text("dont_show_again_btn"), command=on_dont_show).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text=get_text("ok_btn"), command=on_ok).pack(side=tk.RIGHT, padx=5)
        
        top.transient(self.root)
        top.grab_set()
        self.root.wait_window(top)
        
    def show_mineru_info(self):
        info = get_text("mineru_info_content")
        top = self.create_toplevel(get_text("mineru_info_title"), 640, 360)
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
                messagebox.showerror(get_text("error_btn"), f"无法打开网页: {e}")

        ttk.Button(btn_frame, text=get_text("open_mineru_website"), command=open_mineru_website).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_frame, text=get_text("close_btn"), command=top.destroy).pack(side=tk.LEFT, padx=6)
        
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
                    raise ValueError(get_text("offset_value_error"))

            ratio = min(screen_width/16, screen_height/9)
            max_display_width = int(16 * ratio)
            max_display_height = int(9 * ratio)

            display_width = int(max_display_width * self.ratio_var.get())
            display_height = int(max_display_height * self.ratio_var.get())

            print(get_text("start_processing", file=pdf_file))

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
                raise ValueError(get_text("page_range_error"))
            
            # 根据页码范围生成文件名后缀
            page_suffix = format_page_suffix(pages_list)
            out_ppt_file = workspace_dir / f"{pdf_name}{page_suffix}.pptx"
            
            method_id = self.get_method_id_from_translated_name(self.inpaint_method_var.get())

            
            if self.image_only_var.get():
                print("=" * 60)
                print(get_text("image_only_mode_start"))
                print("=" * 60)
                
                png_names = pdf_to_png(
                    pdf_path=pdf_file,
                    output_dir=png_dir,
                    dpi=self.dpi_var.get(),
                    inpaint=self.inpaint_var.get(),
                    pages=pages_list,
                    inpaint_method=method_id,
                    force_regenerate=self.force_regenerate_var.get()
                )
                
                if self.stop_flag:
                    print("\n" + get_text("conversion_stopped_msg"))
                    messagebox.showinfo(get_text("conversion_stopped_title"), get_text("conversion_stopped_msg"))
                    return
                
                png_names = create_ppt_from_images(png_dir, out_ppt_file, png_names=png_names)
                extra_message = f" ({get_text('image_only_label')})"
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
                    inpaint_method=method_id,
                    top_left=self.top_left
                )

                if self.stop_flag:
                    print("\n" + get_text("conversion_stopped_msg"))
                    messagebox.showinfo(get_text("conversion_stopped_title"), get_text("conversion_stopped_msg"))
                    return

                png_names = combine_ppt(ppt_dir, out_ppt_file, png_names=png_names)
                extra_message = ""
            if not self.image_only_var.get():
                # 如果用户提供了 mineru JSON，则进行 refine_ppt 处理
                mineru_json = self.mineru_json_var.get().strip().strip('"')
                if mineru_json:
                    if not os.path.exists(mineru_json):
                        print(f"⚠️ {mineru_json} not exists")
                    else:
                        refined_out = workspace_dir / f"{pdf_name}{page_suffix}_optimized.pptx"
                        print(get_text("mineru_optimizing", file=mineru_json))
                        refine_ppt(str(tmp_image_dir), mineru_json, str(out_ppt_file), str(png_dir), png_names, str(refined_out), unify_font=unify_font, font_name=font_name)
                        
                        print(get_text("refine_ppt_done"))
                        extra_message = "\n\n" + get_text("refine_extra_msg")
                        out_ppt_file = os.path.abspath(refined_out)
                else:
                    extra_message = ""
            else:
                extra_message = f" ({get_text('image_only_label')})"
            out_ppt_file = os.path.abspath(out_ppt_file)
            print("\n" + get_text("conversion_done"))
            print(get_text("output_file", file=out_ppt_file))
            os.startfile(out_ppt_file)
            messagebox.showinfo(get_text("conversion_success_title"), get_text("conversion_success_msg", file=out_ppt_file) + extra_message)
        except Exception as e:
            print("\n" + get_text("conversion_fail", error=str(e)))
            messagebox.showerror(get_text("conversion_fail_title"), get_text("conversion_fail_msg", error=str(e)))
        finally:
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)

    def add_task(self, pdf_path, json_path=None):
        # 捕获当前界面的所有设置
        settings = {
            "output_dir": self.output_dir_var.get().strip().strip('"'),
            "dpi": self.dpi_var.get(),
            "ratio": self.ratio_var.get(),
            "inpaint": self.inpaint_var.get(),
            "inpaint_method": self.inpaint_method_var.get(),
            "image_only": self.image_only_var.get(),
            "force_regenerate": self.force_regenerate_var.get(),
            "unify_font": self.unify_font_var.get(),
            "font_name": self.font_name_var.get().strip() or "Calibri",
            "page_range": self.page_range_var.get().strip()
        }

        # 检查是否已存在相同 PDF 的任务
        for task in self.task_queue:
            if task["pdf"] == pdf_path:
                # 更新现有任务的 JSON 路径和状态，以及所有设置
                task["json"] = json_path or ""
                task["status"] = get_text("queue_status_pending")
                task["settings"] = settings
                self.update_task_row(task)
                print(get_text("queue_task_updated", file=pdf_path))
                return

        task = {
            "id": self.task_id_counter,
            "pdf": pdf_path,
            "json": json_path or "",
            "status": get_text("queue_status_pending"),
            "output": "",
            "settings": settings
        }
        self.task_queue.append(task)
        self.queue_tree.insert("", tk.END, iid=str(task["id"]), values=(task["id"], task["pdf"], task["json"], task["status"], task["output"]))
        self.task_id_counter += 1
        print(get_text("queue_task_added", file=pdf_path))

    def add_task_dialog(self):
        # 优先使用界面上已选择的文件
        current_pdf = self.pdf_path_var.get().strip().strip('"')
        current_json = self.mineru_json_var.get().strip().strip('"')
        
        if current_pdf and os.path.exists(current_pdf):
            # 如果已经选了 PDF，直接添加到队列
            self.add_task(current_pdf, current_json or None)
            # 添加完后可以考虑清空主界面，或者保持不变。通常保持不变比较好，用户可能想改个参数再加一个
            return

        pdf = filedialog.askopenfilename(parent=self.root, title=get_text("select_pdf_title"), filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not pdf:
            return
        json_file = filedialog.askopenfilename(parent=self.root, title=get_text("select_json_title"), filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if json_file and not os.path.exists(json_file):
            json_file = ""
        self.add_task(pdf, json_file or None)

    def add_tasks_multi_pdfs(self):
        pdfs = filedialog.askopenfilenames(parent=self.root, title=get_text("select_pdf_title"), filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not pdfs:
            return
        for p in pdfs:
            self.add_task(p, None)

    def remove_selected_task(self):
        sel = self.queue_tree.selection()
        if not sel:
            return
        for iid in sel:
            tid = int(iid)
            self.queue_tree.delete(iid)
            self.task_queue = [t for t in self.task_queue if t["id"] != tid]
        print(get_text("queue_task_removed"))

    def clear_tasks(self):
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)
        self.task_queue = []
        print(get_text("queue_cleared"))

    def on_task_double_click(self, event):
        item_id = self.queue_tree.identify_row(event.y)
        if not item_id:
            return
        
        # item_id is stored as string of task id
        try:
            task_id = int(item_id)
        except ValueError:
            return

        task = next((t for t in self.task_queue if t["id"] == task_id), None)
        if not task:
            return

        self.show_task_details(task)

    def show_task_details(self, task):
        # 减小默认高度到 650，宽度保持 700，适配更多屏幕
        top = self.create_toplevel(get_text("task_details_title"), 700, 650)
        
        # 使用 Canvas 和 Scrollbar 支持滚动
        canvas = tk.Canvas(top, highlightthickness=0)
        scrollbar = ttk.Scrollbar(top, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # 让内部框架宽度随 Canvas 变化，防止内容被遮挡
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))
        
        canvas.configure(yscrollcommand=scrollbar.set)

        info_frame = ttk.LabelFrame(scrollable_frame, text=get_text("queue_label"), padding="10")
        info_frame.pack(fill=tk.X, expand=True, padx=10, pady=5)
        
        # 存储可编辑的变量
        edit_vars = {}

        # Helper to create rows (只读)
        def add_readonly_row(parent, label_key, value, row, is_path=False):
            # 增加标签宽度到 120，移除换行限制，使布局更舒展
            ttk.Label(parent, text=get_text(label_key), font=("", 9, "bold")).grid(row=row, column=0, sticky="nw", pady=5)
            val_frame = ttk.Frame(parent)
            val_frame.grid(row=row, column=1, sticky="ew", padx=10, pady=5)
            
            display_value = str(value) if value is not None else get_text("none")
            
            # 使用较高的 height 确保长路径显示，width 设置为 10 并配合 expand 填充
            txt = tk.Text(val_frame, height=3 if is_path else 1, width=10, wrap=tk.CHAR, borderwidth=0, bg=top.cget("bg"), font=("", 9))
            txt.insert("1.0", display_value)
            txt.configure(state="disabled")
            txt.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            if is_path and value and os.path.exists(value):
                def open_path():
                    try:
                        if os.path.isfile(value):
                            os.startfile(os.path.dirname(value))
                        else:
                            os.startfile(value)
                    except Exception:
                        pass
                ttk.Button(val_frame, text=get_text("open_btn"), command=open_path, width=6).pack(side=tk.RIGHT, padx=5)

        info_frame.columnconfigure(1, weight=1)
        
        # 判断任务是否可编辑（只有未开始或出错的任务可以修改）
        is_editable = task["status"] not in [get_text("queue_status_done"), get_text("queue_status_running")]
        
        # 统一锁定状态下的颜色：使用 Entry 的 readonly 状态，并设置 readonlybackground 为窗口背景色，确保文字清晰
        # 即使是禁用状态，也要保证文字颜色是黑色或深灰
        style = ttk.Style()
        style.configure("ReadOnly.TEntry", fieldbackground=top.cget("bg"), foreground="black")
        style.configure("ReadOnly.TCombobox", fieldbackground=top.cget("bg"), foreground="black")

        widget_state = "normal" if is_editable else "disabled"
        entry_state = "normal" if is_editable else "readonly"

        add_readonly_row(info_frame, "queue_col_id", task["id"], 0)
        add_readonly_row(info_frame, "queue_col_status", task["status"], 1)
        add_readonly_row(info_frame, "queue_col_pdf", task["pdf"], 2, is_path=True)
        
        # JSON 路径支持编辑
        ttk.Label(info_frame, text=get_text("queue_col_json"), font=("", 9, "bold")).grid(row=3, column=0, sticky="nw", pady=5)
        json_frame = ttk.Frame(info_frame)
        json_frame.grid(row=3, column=1, sticky="ew", padx=10, pady=5)
        json_var = tk.StringVar(value=task["json"] or "")
        ttk.Entry(json_frame, textvariable=json_var, state=entry_state, style="ReadOnly.TEntry" if not is_editable else "").pack(side=tk.LEFT, fill=tk.X, expand=True)
        def browse_json():
            f = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
            if f: json_var.set(f)
        ttk.Button(json_frame, text=get_text("browse_btn"), command=browse_json, width=6, state=widget_state).pack(side=tk.RIGHT, padx=5)

        add_readonly_row(info_frame, "queue_col_output", task["output"], 4, is_path=True)

        # 设置项区域 - 可编辑
        settings = task.get("settings", {})
        # 即使没有settings，也要初始化一个空字典以便编辑（理论上新建任务都有）
        if settings is None:
            settings = {}
            
        set_frame = ttk.LabelFrame(scrollable_frame, text=get_text("task_settings_title"), padding="10")
        set_frame.pack(fill=tk.X, expand=True, padx=10, pady=5)
        set_frame.columnconfigure(1, weight=1)
        
        # 定义每个设置项的类型和对应的键
        # type: entry, int_entry, bool, combo
        s_items = [
            ("output_dir_label", "output_dir", "dir_entry"),
            ("dpi_label", "dpi", "int_entry"),
            ("ratio_label", "ratio", "float_entry"),
            ("inpaint_label", "inpaint", "bool"),
            ("inpaint_method_label", "inpaint_method", "combo_method"),
            ("image_only_label", "image_only", "bool"),
            ("force_regenerate_label", "force_regenerate", "bool"),
            ("unify_font_label", "unify_font", "bool"),
            ("font_name_label", "font_name", "entry"),
            ("page_range_label", "page_range", "entry"),
        ]
        
        # 存储可能需要隐藏的组件
        unify_font_widgets = []
        font_name_widgets = []
        
        for i, (lbl_key, set_key, widget_type) in enumerate(s_items):
            # 移除 wraplength，让标签自然显示，同时增加 grid 的权重配置
            lbl = ttk.Label(set_frame, text=get_text(lbl_key), font=("", 9, "bold"))
            lbl.grid(row=i, column=0, sticky="nw", pady=5)
            
            curr_val = settings.get(set_key)
            widget = None
            
            if widget_type == "bool":
                var = tk.BooleanVar(value=bool(curr_val))
                edit_vars[set_key] = var
                # Checkbutton 在禁用时很难看清，如果是锁定状态，我们用 Label 显示“是/否”
                if not is_editable:
                    val_text = get_text("yes") if var.get() else get_text("no")
                    widget = ttk.Label(set_frame, text=val_text)
                    widget.grid(row=i, column=1, sticky="w", padx=10)
                else:
                    widget = ttk.Checkbutton(set_frame, variable=var)
                    widget.grid(row=i, column=1, sticky="w", padx=10)
                
            elif widget_type == "combo_method":
                var = tk.StringVar(value=str(curr_val))
                edit_vars[set_key] = var
                if not is_editable:
                    # 锁定状态用只读 Entry 模拟，确保文字清晰
                    widget = ttk.Entry(set_frame, textvariable=var, state="readonly", style="ReadOnly.TEntry")
                    widget.grid(row=i, column=1, sticky="ew", padx=10)
                else:
                    widget = ttk.Combobox(set_frame, textvariable=var, values=self.get_translated_method_names(), state="readonly")
                    widget.grid(row=i, column=1, sticky="ew", padx=10)
                
            elif widget_type == "dir_entry":
                widget = ttk.Frame(set_frame)
                widget.grid(row=i, column=1, sticky="ew", padx=10)
                var = tk.StringVar(value=str(curr_val) if curr_val is not None else "")
                edit_vars[set_key] = var
                ttk.Entry(widget, textvariable=var, state=entry_state, style="ReadOnly.TEntry" if not is_editable else "").pack(side=tk.LEFT, fill=tk.X, expand=True)
                def browse_dir(v=var):
                    d = filedialog.askdirectory(parent=top)
                    if d: v.set(d)
                if is_editable:
                    ttk.Button(widget, text=get_text("browse_btn"), command=browse_dir).pack(side=tk.RIGHT, padx=5)
                
            else: # entry, int_entry, float_entry
                var = tk.StringVar(value=str(curr_val) if curr_val is not None else "")
                edit_vars[set_key] = var
                widget = ttk.Entry(set_frame, textvariable=var, state=entry_state, style="ReadOnly.TEntry" if not is_editable else "")
                widget.grid(row=i, column=1, sticky="ew", padx=10)

            if set_key == "unify_font":
                unify_font_widgets = [lbl, widget, i]
            elif set_key == "font_name":
                font_name_widgets = [lbl, widget, i]

        def update_details_unify_font_visibility(*args):
            is_visible = bool(json_var.get().strip())
            if is_visible:
                unify_font_widgets[0].grid(row=unify_font_widgets[2], column=0, sticky="nw", pady=5)
                unify_font_widgets[1].grid(row=unify_font_widgets[2], column=1, sticky="w", padx=10)
                
                font_name_widgets[0].grid(row=font_name_widgets[2], column=0, sticky="nw", pady=5)
                font_name_widgets[1].grid(row=font_name_widgets[2], column=1, sticky="ew", padx=10)
            else:
                unify_font_widgets[0].grid_forget()
                unify_font_widgets[1].grid_forget()
                font_name_widgets[0].grid_forget()
                font_name_widgets[1].grid_forget()

        json_var.trace_add("write", update_details_unify_font_visibility)
        update_details_unify_font_visibility()

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 底部按钮区域
        btn_frame = ttk.Frame(top, padding="10")
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        def save_changes():
            # 更新 task["settings"]
            new_settings = {}
            for k, v in edit_vars.items():
                val = v.get()
                # 类型转换处理
                # 注意：这里为了简单起见，大部分存为原始类型，读取时 run_conversion_for_task 会再处理
                # 但为了保持一致性，尽量还原类型
                orig_type_item = next((x for x in s_items if x[1] == k), None)
                if orig_type_item:
                    w_type = orig_type_item[2]
                    if w_type == "int_entry":
                        try:
                            new_settings[k] = int(val)
                        except:
                            new_settings[k] = val # fallback
                    elif w_type == "float_entry":
                        try:
                            new_settings[k] = float(val)
                        except:
                            new_settings[k] = val
                    elif w_type == "bool":
                        new_settings[k] = bool(val)
                    else:
                        new_settings[k] = val
                else:
                    new_settings[k] = val
            
            task["settings"] = new_settings
            # 更新 JSON 路径
            task["json"] = json_var.get().strip().strip('"')
            # 更新 Treeview 显示
            self.update_task_row(task)
            
            print(f"Task {task['id']} updated.")
            top.destroy()

        ttk.Button(btn_frame, text=get_text("close_btn"), command=top.destroy).pack(side=tk.RIGHT, padx=5)
        # 只有未开始的任务允许修改设置
        if is_editable:
            ttk.Button(btn_frame, text=get_text("save_btn"), command=save_changes).pack(side=tk.RIGHT, padx=5)

    def update_task_row(self, task):
        self.queue_tree.item(str(task["id"]), values=(task["id"], task["pdf"], task["json"], task["status"], task["output"]))

    def start_queue(self):
        if self.is_queue_running:
            return

        # 如果队列为空，但主界面有选 PDF，则先自动添加
        if not self.task_queue:
            current_pdf = self.pdf_path_var.get().strip().strip('"')
            current_json = self.mineru_json_var.get().strip().strip('"')
            if current_pdf and os.path.exists(current_pdf):
                self.add_task(current_pdf, current_json or None)

        if not self.image_only_var.get():
            if not self.ensure_pc_manager_running():
                return
        if not self.task_queue:
            messagebox.showinfo(get_text("info_btn"), get_text("queue_empty_msg"))
            return
        self.queue_stop_flag = False
        self.is_queue_running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        threading.Thread(target=self.process_queue, daemon=True).start()

    def stop_queue(self):
        if not self.is_queue_running:
            return
        self.queue_stop_flag = True
        print(get_text("queue_stopping"))
        self.stop_btn.config(state=tk.DISABLED)

    def process_queue(self):
        try:
            print(get_text("queue_started"))
            for task in list(self.task_queue):
                if self.queue_stop_flag:
                    break
                if not os.path.exists(task["pdf"]):
                    task["status"] = get_text("queue_status_error")
                    self.update_task_row(task)
                    continue
                task["status"] = get_text("queue_status_running")
                self.update_task_row(task)
                ok, out_file = self.run_conversion_for_task(task)
                task["output"] = out_file or ""
                task["status"] = get_text("queue_status_done") if ok else get_text("queue_status_error")
                self.update_task_row(task)
            if self.queue_stop_flag:
                print(get_text("queue_stopped"))
            else:
                print(get_text("queue_finished"))
        finally:
            self.is_queue_running = False
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)

    def run_conversion_for_task(self, task):
        try:
            pdf_file = task["pdf"]
            mineru_json = task["json"]
            settings = task.get("settings", {})
            
            # 从任务设置中提取参数，如果缺失则使用当前界面值（兼容旧任务）
            output_dir = settings.get("output_dir", self.output_dir_var.get())
            dpi = settings.get("dpi", self.dpi_var.get())
            ratio_val = settings.get("ratio", self.ratio_var.get())
            inpaint = settings.get("inpaint", self.inpaint_var.get())
            inpaint_method = settings.get("inpaint_method", self.inpaint_method_var.get())
            image_only = settings.get("image_only", self.image_only_var.get())
            force_regenerate = settings.get("force_regenerate", self.force_regenerate_var.get())
            unify_font = settings.get("unify_font", self.unify_font_var.get())
            font_name = settings.get("font_name", self.font_name_var.get())
            page_range = settings.get("page_range", self.page_range_var.get())
            
            # 全局设置（不随任务存储，始终使用界面当前值）
            delay = self.delay_var.get()
            timeout = self.timeout_var.get()
            done_offset_str = self.done_offset_var.get().strip()
            calibrate = self.calibrate_var.get()

            pdf_name = Path(pdf_file).stem
            workspace_dir = Path(output_dir)
            png_dir = workspace_dir / f"{pdf_name}_pngs"
            ppt_dir = workspace_dir / f"{pdf_name}_ppt"
            tmp_image_dir = workspace_dir / "tmp_images"
            workspace_dir.mkdir(exist_ok=True, parents=True)
            
            done_offset = None
            if done_offset_str:
                try:
                    done_offset = int(done_offset_str)
                except ValueError:
                    pass

            ratio = min(screen_width/16, screen_height/9)
            max_display_width = int(16 * ratio)
            max_display_height = int(9 * ratio)
            display_width = int(max_display_width * ratio_val)
            display_height = int(max_display_height * ratio_val)
            
            def parse_page_range(range_str):
                if not range_str:
                    return None
                pages = set()
                range_str = range_str.replace('，', ',')
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
                pages_list = parse_page_range(page_range)
            except Exception:
                pages_list = None
                
            page_suffix = format_page_suffix(pages_list)
            out_ppt_file = workspace_dir / f"{pdf_name}{page_suffix}.pptx"
            
            method_id = self.get_method_id_from_translated_name(inpaint_method)
            
            if image_only:
                png_names = pdf_to_png(
                    pdf_path=pdf_file,
                    output_dir=png_dir,
                    dpi=dpi,
                    inpaint=inpaint,
                    pages=pages_list,
                    inpaint_method=method_id,
                    force_regenerate=force_regenerate
                )
                if self.queue_stop_flag:
                    return False, None
                png_names = create_ppt_from_images(png_dir, out_ppt_file, png_names=png_names)
            else:
                png_names = process_pdf_to_ppt(
                    pdf_path=pdf_file,
                    png_dir=png_dir,
                    ppt_dir=ppt_dir,
                    delay_between_images=delay,
                    inpaint=inpaint,
                    dpi=dpi,
                    timeout=timeout,
                    display_height=display_height,
                    display_width=display_width,
                    done_button_offset=done_offset,
                    capture_done_offset=calibrate,
                    pages=pages_list,
                    update_offset_callback=self.update_offset_disk,
                    stop_flag=lambda: self.queue_stop_flag,
                    force_regenerate=force_regenerate,
                    inpaint_method=method_id,
                    top_left=self.top_left
                )
                if self.queue_stop_flag:
                    return False, None
                png_names = combine_ppt(ppt_dir, out_ppt_file, png_names=png_names)
                
            if not image_only and mineru_json:
                if os.path.exists(mineru_json):
                    refined_out = workspace_dir / f"{pdf_name}{page_suffix}_optimized.pptx"
                    refine_ppt(str(tmp_image_dir), mineru_json, str(out_ppt_file), str(png_dir), png_names, str(refined_out), unify_font=unify_font)
                    out_ppt_file = os.path.abspath(refined_out)
            
            out_ppt_file = os.path.abspath(out_ppt_file)
            print(get_text("queue_task_done", file=out_ppt_file))
            return True, out_ppt_file
        except Exception as e:
            print(get_text("conversion_fail", error=str(e)))
            return False, None

def launch_gui():
    # Enable Windows DPI awareness before creating the Tk root where possible
    try:
        enable_windows_dpi_awareness(None)
    except Exception:
        pass

    root = tk.Tk()
    # 设置窗口图标
    root.iconbitmap(icon_path())
    # After root exists, apply scaling using the helper (this will call tk scaling)
    try:
        enable_windows_dpi_awareness(root)
    except Exception:
        pass

    app = AppGUI(root)
    root.mainloop()

if __name__ == "__main__":
    launch_gui()
