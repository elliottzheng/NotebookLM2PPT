import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import sys
import os
try:
    import windnd
except ImportError:
    print("windnd 模块未安装，拖拽功能将不可用。")
    windnd = None
from pathlib import Path
from .cli import process_pdf_to_ppt
from .ppt_combiner import combine_ppt
from .utils.screenshot_automation import screen_width, screen_height, load_saved_done_offset

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
        self.root.title("PDF to PPT Converter")
        self.root.geometry("800x600")
        self.root.minsize(700, 480)
        
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
            # Get the first file dropped
            file_path = files[0].decode('gbk') if isinstance(files[0], bytes) else files[0]
            if file_path.lower().endswith('.pdf'):
                self.pdf_path_var.set(file_path)
                print(f"已通过拖拽选择文件: {file_path}")
            else:
                messagebox.showwarning("警告", "只支持 PDF 文件")

    def on_closing(self):
        # Restore stdout/stderr
        sys.stdout = self.old_stdout
        sys.stderr = self.old_stderr
        self.root.destroy()

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
        file_frame = ttk.LabelFrame(main_frame, text="文件设置 (支持拖拽 PDF 文件到窗口)", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="PDF 文件:").grid(row=0, column=0, sticky=tk.W)
        self.pdf_path_var = tk.StringVar()
        pdf_entry = ttk.Entry(file_frame, textvariable=self.pdf_path_var, width=60)
        pdf_entry.grid(row=0, column=1, padx=5, sticky="ew")
        self.add_context_menu(pdf_entry)
        ttk.Button(file_frame, text="浏览", command=self.browse_pdf).grid(row=0, column=2)

        ttk.Label(file_frame, text="输出目录:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_dir_var = tk.StringVar(value="workspace")
        output_entry = ttk.Entry(file_frame, textvariable=self.output_dir_var, width=60)
        output_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.add_context_menu(output_entry)
        ttk.Button(file_frame, text="浏览", command=self.browse_output).grid(row=1, column=2, pady=5)

        # Options
        opt_frame = ttk.LabelFrame(main_frame, text="转换选项", padding="10")
        opt_frame.pack(fill=tk.X, pady=5)
        opt_frame.columnconfigure(1, weight=1)
        opt_frame.columnconfigure(3, weight=1)

        ttk.Label(opt_frame, text="DPI:").grid(row=0, column=0, sticky=tk.W)
        self.dpi_var = tk.IntVar(value=150)
        dpi_entry = ttk.Entry(opt_frame, textvariable=self.dpi_var, width=10)
        dpi_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.add_context_menu(dpi_entry)

        ttk.Label(opt_frame, text="延迟 (秒):").grid(row=0, column=2, sticky=tk.W, padx=10)
        self.delay_var = tk.IntVar(value=2)
        delay_entry = ttk.Entry(opt_frame, textvariable=self.delay_var, width=10)
        delay_entry.grid(row=0, column=3, sticky=tk.W, padx=5)
        self.add_context_menu(delay_entry)

        ttk.Label(opt_frame, text="超时 (秒):").grid(row=0, column=4, sticky=tk.W, padx=10)
        self.timeout_var = tk.IntVar(value=50)
        timeout_entry = ttk.Entry(opt_frame, textvariable=self.timeout_var, width=10)
        timeout_entry.grid(row=0, column=5, sticky=tk.W, padx=5)
        self.add_context_menu(timeout_entry)

        ttk.Label(opt_frame, text="显示比例:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.ratio_var = tk.DoubleVar(value=0.8)
        ratio_entry = ttk.Entry(opt_frame, textvariable=self.ratio_var, width=10)
        ratio_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        self.add_context_menu(ratio_entry)

        # 将页范围放到单独的行，避免与偏移提示重叠
        ttk.Label(opt_frame, text="页范围:").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        self.page_range_var = tk.StringVar(value="")
        page_range_entry = ttk.Entry(opt_frame, textvariable=self.page_range_var, width=30)
        page_range_entry.grid(row=5, column=1, columnspan=3, sticky="ew", padx=5, pady=5)
        self.add_context_menu(page_range_entry)
        ttk.Label(opt_frame, text="示例: 1-3,5,7- (与 Word 打印页范围一致)", wraplength=420).grid(row=6, column=0, columnspan=4, sticky=tk.W)

        self.inpaint_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="启用图像修复 (去水印)", variable=self.inpaint_var).grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=10)


        ttk.Label(opt_frame, text="转换按钮偏移:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.done_offset_var = tk.StringVar(value="")
        done_offset_entry = ttk.Entry(opt_frame, textvariable=self.done_offset_var, width=10)
        done_offset_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        self.add_context_menu(done_offset_entry)
        # 显示已保存的偏移（不作为手动覆盖输入）
        self.saved_offset_var = tk.StringVar(value="")
        ttk.Label(opt_frame, textvariable=self.saved_offset_var).grid(row=2, column=2, sticky=tk.W, padx=5)
        # 将长说明放到独立一行，跨越所有列并允许横向扩展与自动换行
        ttk.Label(opt_frame, text="该数值表示从右下角到转换按钮的像素偏移(从右往左)，留空则在无已保存偏移时强制按钮位置校准；填数字将作为手动覆盖。", wraplength=640).grid(row=3, column=0, columnspan=6, sticky="ew", pady=2)

        # 首次校准选项：允许用户在第一页手动点击完成按钮以捕获偏移并保存
        # 如果磁盘已有保存偏移，默认关闭首次校准；否则默认开启
        self.calibrate_var = tk.BooleanVar(value=True)

        ttk.Checkbutton(opt_frame, text="按钮位置校准（若无已保存偏移则默认开启）", variable=self.calibrate_var).grid(row=4, column=0, columnspan=4, sticky=tk.W, pady=5)
        self.load_offset_from_disk()

        # Control
        ctrl_frame = ttk.Frame(main_frame, padding="10")
        ctrl_frame.pack(fill=tk.X)

        self.start_btn = ttk.Button(ctrl_frame, text="开始转换", command=self.start_conversion)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        # Log Area
        log_frame = ttk.LabelFrame(main_frame, text="日志输出", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled', height=15)
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.tag_config("stderr", foreground="red")

    def browse_pdf(self):
        # 清理路径中的引号和空格，方便用户直接粘贴带引号的路径
        current_path = self.pdf_path_var.get().strip().strip('"')
        initial_dir = os.path.dirname(current_path) if current_path and os.path.exists(os.path.dirname(current_path)) else None
        
        filename = filedialog.askopenfilename(
            parent=self.root,
            title="选择 PDF 文件",
            initialdir=initial_dir,
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.pdf_path_var.set(filename)

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

    def start_conversion(self):
        pdf_path = self.pdf_path_var.get().strip().strip('"')
        output_dir = self.output_dir_var.get().strip().strip('"')
        
        # Update variables with sanitized paths
        self.pdf_path_var.set(pdf_path)
        self.output_dir_var.set(output_dir)

        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("错误", "请选择有效的 PDF 文件")
            return

        self.start_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.run_conversion, daemon=True).start()

    def load_offset_from_disk(self):
        # 从磁盘加载已保存的偏移并更新显示
        try:
            saved = load_saved_done_offset()
        except Exception:
            saved = None
        # 如果磁盘已有保存偏移，显示其值并将偏移预填到输入框；否则提示将要求首次校准
        default_calibrate = True if saved is None else False
        if saved is not None:
            self.saved_offset_var.set(f"已保存偏移: {saved} 像素")
            # 直接将已保存偏移填入偏移输入框，便于用户微调
            if not self.done_offset_var.get().strip():
                self.done_offset_var.set(str(saved))
        else:
            self.saved_offset_var.set("未保存偏移：运行将要求进行按钮位置校准")
        self.calibrate_var.set(default_calibrate)
        

    def run_conversion(self):
        try:
            pdf_file = self.pdf_path_var.get()
            pdf_name = Path(pdf_file).stem
            workspace_dir = Path(self.output_dir_var.get())
            png_dir = workspace_dir / f"{pdf_name}_pngs"
            ppt_dir = workspace_dir / f"{pdf_name}_ppt"
            out_ppt_file = workspace_dir / f"{pdf_name}.pptx"
            
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

            pages_list = None
            try:
                pages_list = parse_page_range(self.page_range_var.get().strip())
            except Exception as e:
                raise ValueError("页范围格式错误，请使用 1-3,5,7- 类似格式")
            
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
                update_offset_callback=self.load_offset_from_disk
            )

            combine_ppt(ppt_dir, out_ppt_file, png_names=png_names)
            out_ppt_file = os.path.abspath(out_ppt_file)
            print(f"\n转换完成！最终文件: {out_ppt_file}")
            # 打开该文件
            os.startfile(out_ppt_file)
            messagebox.showinfo("成功", f"转换完成！\n文件保存至: {out_ppt_file}")
        except Exception as e:
            print(f"\n发生错误: {str(e)}")
            messagebox.showerror("错误", f"转换过程中发生错误: {str(e)}")
        finally:
            self.start_btn.config(state=tk.NORMAL)

def launch_gui():
    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()

if __name__ == "__main__":
    launch_gui()
