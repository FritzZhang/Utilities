"""
PDF2Gray - 跨平台 PDF 转灰度批量工具

使用说明：
1. 安装依赖：
"""
"""
   # requirements.txt 内容：PyMuPDF>=1.24.0
2. 安装 Ghostscript（推荐，提升质量与体积）：
   - Windows: https://www.ghostscript.com/download/gsdnld.html
   - macOS: brew install ghostscript
   - Linux: sudo apt install ghostscript
3. 运行：
   python pdf_to_grayscale_gui.py
4. 打包为单文件可执行：
   pyinstaller --noconsole --onefile --name PDF2Gray pdf_to_grayscale_gui.py

常见问题：
- 找不到 Ghostscript：请检查环境变量或手动指定 gs 路径。
- 输出体积变大：PyMuPDF 回退模式会将页面转为灰度图片，体积可能增大。
- 权限/磁盘空间不足：请检查输出目录权限与剩余空间。

"""
import queue
import json
import logging
import shutil
import platform
import subprocess
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any
from concurrent.futures import ThreadPoolExecutor, Future
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os   

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

APP_NAME = "PDF2Gray"
CONFIG_FILE = str(Path.home() / f".{APP_NAME.lower()}_config.json")
DEFAULT_SUFFIX = "_gray"
DEFAULT_DPI = 200

# ========== 日志设置 ==========
class TkinterLogHandler(logging.Handler):
    def __init__(self, text_widget: scrolledtext.ScrolledText):
        super().__init__()
        self.text_widget = text_widget
        self.text_widget.config(state=tk.NORMAL)
        self.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', "%H:%M:%S"))

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
        self.text_widget.after(0, append)

# ========== 转换核心 ==========
class PDFConverter:
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.gs_path = self.detect_ghostscript()
        self.has_gs = self.gs_path is not None
        if not self.has_gs and fitz is None:
            raise RuntimeError("未检测到 Ghostscript，且未安装 PyMuPDF。请先安装 PyMuPDF: pip install PyMuPDF")

    @staticmethod
    def detect_ghostscript() -> Optional[str]:
        """自动检测 Ghostscript 可执行文件路径"""
        candidates = []
        if platform.system() == "Windows":
            candidates = ["gswin64c.exe", "gswin32c.exe", "gs.exe"]
        else:
            candidates = ["gs"]
        for exe in candidates:
            path = shutil.which(exe)
            if path:
                return path
        return None

    def convert(self, in_pdf: Path, out_pdf: Path,
                gs_args: str = "", fallback_dpi: int = DEFAULT_DPI) -> Tuple[bool, str]:
        """主转换入口，优先 Ghostscript，失败则 PyMuPDF"""
        if self.has_gs:
            ok, msg = self.convert_with_gs(in_pdf, out_pdf, gs_args)
            if ok:
                return True, f"Ghostscript 成功: {msg}"
            else:
                self.logger.warning(f"Ghostscript 失败: {msg}，尝试 PyMuPDF 回退。")
        if fitz is not None:
            ok, msg = self.convert_with_pymupdf(in_pdf, out_pdf, fallback_dpi)
            if ok:
                return True, f"PyMuPDF 成功: {msg}"
            else:
                return False, f"PyMuPDF 失败: {msg}"
        return False, "无可用后端"

    def convert_with_gs(self, in_pdf: Path, out_pdf: Path, gs_args: str = "") -> Tuple[bool, str]:
        """Ghostscript 转换"""
        cmd = [self.gs_path,
               "-sDEVICE=pdfwrite",
               "-sColorConversionStrategy=Gray",
               "-dProcessColorModel=/DeviceGray",
               "-dCompatibilityLevel=1.4",
               "-dDetectDuplicateImages=true",
               "-dDownsampleColorImages=true",
               "-dColorImageResolution=300",
               "-dNOPAUSE",
               "-dBATCH",
               "-dQUIET",
               f"-sOutputFile={str(out_pdf)}",
               str(in_pdf)]
        if gs_args:
            cmd[1:1] = gs_args.strip().split()
        try:
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
            return True, ""
        except Exception as e:
            return False, str(e)

    def convert_with_pymupdf(self, in_pdf: Path, out_pdf: Path, dpi: int) -> Tuple[bool, str]:
        """PyMuPDF 转换（每页渲染为灰度图片）"""
        if fitz is None:
            return False, "未安装 PyMuPDF"
        try:
            doc = fitz.open(str(in_pdf))
            out_doc = fitz.open()
            for i in range(doc.page_count):
                try:
                    page = doc.load_page(i)
                    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY, alpha=False)
                    rect = page.rect
                    pdf_bytes = fitz.open("pdf", pix.tobytes("pdf"))
                    out_doc.insert_pdf(pdf_bytes)
                except Exception as e:
                    self.logger.error(f"第{i+1}页转换失败: {e}")
            out_doc.save(str(out_pdf))
            out_doc.close()
            doc.close()
            return True, ""
        except Exception as e:
            return False, str(e)

# ========== GUI 主体 ==========
class PDF2GrayApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.logger = logging.getLogger(APP_NAME)
        self.logger.setLevel(logging.INFO)
        self.file_list: List[Dict[str, Any]] = []  # {path, size, status, future}
        self.executor: Optional[ThreadPoolExecutor] = None
        self.cancel_flag = threading.Event()
        self.config = self.load_config()
        self.converter = PDFConverter(self.logger)
        self.setup_gui()
        self.refresh_list()

    def load_config(self) -> dict:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {
            "last_input": str(Path.home()),
            "last_output": str(Path.home()),
            "suffix": DEFAULT_SUFFIX,
            "overwrite": False,
            "gs_args": "",
            "dpi": DEFAULT_DPI,
            "threads": os.cpu_count() or 4
        }

    def save_config(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def setup_gui(self):
        # 文件列表区
        frm_top = ttk.Frame(self.root)
        frm_top.pack(fill=tk.BOTH, expand=False, padx=8, pady=4)
        btn_add = ttk.Button(frm_top, text="添加 PDF...", command=self.add_files)
        btn_add.pack(side=tk.LEFT, padx=2)
        btn_add_dir = ttk.Button(frm_top, text="添加文件夹...", command=self.add_folder)
        btn_add_dir.pack(side=tk.LEFT, padx=2)
        btn_remove = ttk.Button(frm_top, text="移除选中", command=self.remove_selected)
        btn_remove.pack(side=tk.LEFT, padx=2)
        btn_clear = ttk.Button(frm_top, text="清空列表", command=self.clear_list)
        btn_clear.pack(side=tk.LEFT, padx=2)
        btn_output = ttk.Button(frm_top, text="选择输出文件夹...", command=self.choose_output_dir)
        btn_output.pack(side=tk.LEFT, padx=2)

        # 选项区
        frm_opts = ttk.LabelFrame(self.root, text="选项")
        frm_opts.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(frm_opts, text="输出命名后缀:").grid(row=0, column=0, sticky=tk.W, padx=2, pady=2)
        self.var_suffix = tk.StringVar(value=self.config.get("suffix", DEFAULT_SUFFIX))
        ent_suffix = ttk.Entry(frm_opts, textvariable=self.var_suffix, width=10)
        ent_suffix.grid(row=0, column=1, padx=2, pady=2)
        ttk.Label(frm_opts, text="回退渲染 DPI:").grid(row=0, column=2, sticky=tk.W, padx=2, pady=2)
        self.var_dpi = tk.IntVar(value=self.config.get("dpi", DEFAULT_DPI))
        ent_dpi = ttk.Entry(frm_opts, textvariable=self.var_dpi, width=6)
        ent_dpi.grid(row=0, column=3, padx=2, pady=2)
        self.var_overwrite = tk.BooleanVar(value=self.config.get("overwrite", False))
        chk_overwrite = ttk.Checkbutton(frm_opts, text="允许覆盖同名文件", variable=self.var_overwrite)
        chk_overwrite.grid(row=0, column=4, padx=2, pady=2)
        # 高级参数
        self.var_gs_args = tk.StringVar(value=self.config.get("gs_args", ""))
        self.adv_frame = ttk.LabelFrame(self.root, text="Ghostscript 参数（高级）")
        self.adv_frame.pack(fill=tk.X, padx=8, pady=2)
        ent_gs_args = ttk.Entry(self.adv_frame, textvariable=self.var_gs_args, width=60)
        ent_gs_args.pack(side=tk.LEFT, padx=2, pady=2)
        self.adv_frame.pack_forget()  # 默认隐藏
        btn_adv = ttk.Button(frm_opts, text="高级参数...", command=self.toggle_adv)
        btn_adv.grid(row=0, column=5, padx=2, pady=2)

        # 文件列表 Treeview
        frm_list = ttk.Frame(self.root)
        frm_list.pack(fill=tk.BOTH, expand=True, padx=8, pady=2)
        columns = ("path", "size", "status")
        self.tree = ttk.Treeview(frm_list, columns=columns, show="headings", selectmode="extended", height=10)
        self.tree.heading("path", text="PDF 路径")
        self.tree.heading("size", text="大小 (MB)")
        self.tree.heading("status", text="状态")
        self.tree.column("path", width=400)
        self.tree.column("size", width=80, anchor=tk.E)
        self.tree.column("status", width=100, anchor=tk.CENTER)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(frm_list, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # 进度条
        frm_prog = ttk.Frame(self.root)
        frm_prog.pack(fill=tk.X, padx=8, pady=2)
        self.pb_total = ttk.Progressbar(frm_prog, orient="horizontal", length=200, mode="determinate")
        self.pb_total.pack(side=tk.LEFT, padx=2, pady=2, fill=tk.X, expand=True)
        self.pb_file = ttk.Progressbar(frm_prog, orient="horizontal", length=200, mode="determinate")
        self.pb_file.pack(side=tk.LEFT, padx=2, pady=2, fill=tk.X, expand=True)

        # 日志区
        frm_log = ttk.LabelFrame(self.root, text="日志")
        frm_log.pack(fill=tk.BOTH, expand=True, padx=8, pady=2)
        self.txt_log = scrolledtext.ScrolledText(frm_log, height=8, state=tk.NORMAL, font=("Consolas", 10))
        self.txt_log.pack(fill=tk.BOTH, expand=True)
        log_handler = TkinterLogHandler(self.txt_log)
        self.logger.addHandler(log_handler)

        # 按钮区
        frm_btn = ttk.Frame(self.root)
        frm_btn.pack(fill=tk.X, padx=8, pady=4)
        self.btn_start = ttk.Button(frm_btn, text="开始转换", command=self.start_convert)
        self.btn_start.pack(side=tk.LEFT, padx=2)
        self.btn_cancel = ttk.Button(frm_btn, text="取消", command=self.cancel_convert, state=tk.DISABLED)
        self.btn_cancel.pack(side=tk.LEFT, padx=2)

        # 输出目录
        self.var_output_dir = tk.StringVar(value=self.config.get("last_output", str(Path.home())))
        self._output_dir_user_set = False  # 标记用户是否手动设置过输出目录

    def toggle_adv(self):
        if self.adv_frame.winfo_ismapped():
            self.adv_frame.pack_forget()
        else:
            self.adv_frame.pack(fill=tk.X, padx=8, pady=2)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf")],
            initialdir=self.config.get("last_input", str(Path.home()))
        )
        if files:
            self.config["last_input"] = str(Path(files[0]).parent)
            # 如果用户未手动设置过输出目录，则自动设为第一个文件的文件夹
            if not self._output_dir_user_set:
                self.var_output_dir.set(str(Path(files[0]).parent))
            self.save_config()
            for f in files:
                self._add_file(Path(f))
            self.refresh_list()

    def add_folder(self):
        folder = filedialog.askdirectory(
            title="选择文件夹",
            initialdir=self.config.get("last_input", str(Path.home()))
        )
        if folder:
            self.config["last_input"] = folder
            # 如果用户未手动设置过输出目录，则自动设为该文件夹
            if not self._output_dir_user_set:
                self.var_output_dir.set(folder)
            self.save_config()
            pdfs = list(Path(folder).rglob("*.pdf"))
            for f in pdfs:
                self._add_file(f)
            self.refresh_list()

    def _add_file(self, f: Path):
        f = f.resolve()
        if not f.exists() or not f.is_file() or f.suffix.lower() != ".pdf":
            return
        if any(str(f) == item["path"] for item in self.file_list):
            return
        size_mb = f.stat().st_size / 1024 / 1024
        self.file_list.append({
            "path": str(f),
            "size": f"{size_mb:.2f}",
            "status": "待处理",
            "future": None
        })

    def remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        sel_paths = [self.tree.item(i, "values")[0] for i in sel]
        self.file_list = [item for item in self.file_list if item["path"] not in sel_paths]
        self.refresh_list()

    def clear_list(self):
        self.file_list.clear()
        self.refresh_list()

    def choose_output_dir(self):
        d = filedialog.askdirectory(
            title="选择输出文件夹",
            initialdir=self.var_output_dir.get()
        )
        if d:
            self.var_output_dir.set(d)
            self.config["last_output"] = d
            self._output_dir_user_set = True
            self.save_config()

    def refresh_list(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for item in self.file_list:
            self.tree.insert("", tk.END, values=(item["path"], item["size"], item["status"]))

    def start_convert(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加 PDF 文件！")
            return
        out_dir = Path(self.var_output_dir.get()).resolve()
        if not out_dir.exists():
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录: {e}")
                return
        self.config["suffix"] = self.var_suffix.get()
        self.config["overwrite"] = self.var_overwrite.get()
        self.config["gs_args"] = self.var_gs_args.get()
        self.config["dpi"] = self.var_dpi.get()
        self.save_config()
        self.cancel_flag.clear()
        self.btn_start.config(state=tk.DISABLED)
        self.btn_cancel.config(state=tk.NORMAL)
        self.pb_total.config(maximum=len(self.file_list), value=0)
        self.pb_file.config(maximum=100, value=0)
        for item in self.file_list:
            item["status"] = "待处理"
            item["future"] = None
        self.refresh_list()
        self.logger.info(f"开始批量转换，共 {len(self.file_list)} 个 PDF。")
        self.executor = ThreadPoolExecutor(max_workers=self.config.get("threads", os.cpu_count() or 4))
        self.futures: List[Future] = []
        self.success = 0
        self.failed = 0
        self.skipped = 0
        self.total = len(self.file_list)
        self._run_batch(out_dir)

    def _run_batch(self, out_dir: Path):
        def task_wrapper(idx: int, item: dict):
            if self.cancel_flag.is_set():
                return "跳过"
            in_pdf = Path(item["path"])
            out_pdf = out_dir / (in_pdf.stem + self.var_suffix.get() + ".pdf")
            if (not self.var_overwrite.get()) and out_pdf.exists():
                self.logger.info(f"跳过已存在: {out_pdf}")
                return "跳过"
            try:
                ok, msg = self.converter.convert(
                    in_pdf, out_pdf,
                    gs_args=self.var_gs_args.get(),
                    fallback_dpi=self.var_dpi.get()
                )
                if ok:
                    self.logger.info(f"[{in_pdf.name}] 转换成功 ({'GS' if self.converter.has_gs else 'PyMuPDF'}) -> {out_pdf}")
                    return "成功"
                else:
                    self.logger.error(f"[{in_pdf.name}] 转换失败: {msg}")
                    return "失败"
            except Exception as e:
                self.logger.error(f"[{in_pdf.name}] 异常: {e}")
                return "失败"

        def update_status(idx: int, status: str):
            self.file_list[idx]["status"] = status
            self.refresh_list()
            self.pb_total.step(1)
            if status == "成功":
                self.success += 1
            elif status == "失败":
                self.failed += 1
            elif status == "跳过":
                self.skipped += 1
            self.root.update_idletasks()

        def run_all():
            for idx, item in enumerate(self.file_list):
                if self.cancel_flag.is_set():
                    update_status(idx, "跳过")
                    continue
                update_status(idx, "处理中")
                status = task_wrapper(idx, item)
                update_status(idx, status)
                self.pb_file.config(value=100)
            self.logger.info(f"全部完成：成功 {self.success}，失败 {self.failed}，跳过 {self.skipped}")
            self.btn_start.config(state=tk.NORMAL)
            self.btn_cancel.config(state=tk.DISABLED)

        threading.Thread(target=run_all, daemon=True).start()

    def cancel_convert(self):
        self.cancel_flag.set()
        self.logger.warning("用户取消，剩余任务将跳过。")
        self.btn_cancel.config(state=tk.DISABLED)

# ========== 主入口 ==========
def main():
    root = tk.Tk()
    app = PDF2GrayApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
