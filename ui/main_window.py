from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core.processor import ExcelFlowProcessor, ProcessingError, ProcessStats


class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("成交流水处理工具")
        self.root.geometry("900x650")

        self.folder1_var = tk.StringVar()
        self.folder2_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.output_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        input_group = ttk.LabelFrame(main_frame, text="输入参数", padding=10)
        input_group.pack(fill=tk.X)

        self._add_picker_row(input_group, "成交信息 pb 文件夹", self.folder1_var, self._pick_folder)
        self._add_picker_row(input_group, "协议回购成交信息 pb 文件夹", self.folder2_var, self._pick_folder)
        self._add_picker_row(input_group, "Excel 模板文件", self.template_var, self._pick_template)
        self._add_picker_row(input_group, "输出文件路径", self.output_var, self._pick_output_file)

        action_group = ttk.Frame(main_frame, padding=(0, 12))
        action_group.pack(fill=tk.X)

        self.start_btn = ttk.Button(action_group, text="开始处理", command=self.start_processing)
        self.start_btn.pack(side=tk.LEFT)

        ttk.Button(action_group, text="重置", command=self.reset_inputs).pack(side=tk.LEFT, padx=(8, 0))

        log_group = ttk.LabelFrame(main_frame, text="处理日志", padding=10)
        log_group.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_group, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(log_group, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _add_picker_row(self, parent, label, var, pick_cmd):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=4)

        ttk.Label(row, text=label, width=26).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(row, text="选择", command=lambda v=var: pick_cmd(v)).pack(side=tk.LEFT)

    def _pick_folder(self, var):
        selected = filedialog.askdirectory()
        if selected:
            var.set(selected)

    def _pick_template(self, var):
        selected = filedialog.askopenfilename(
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if selected:
            var.set(selected)

    def _pick_output_file(self, var):
        selected = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if selected:
            var.set(selected)

    def log(self, message: str):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def reset_inputs(self):
        self.folder1_var.set("")
        self.folder2_var.set("")
        self.template_var.set("")
        self.output_var.set("")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def start_processing(self):
        folder1 = self.folder1_var.get().strip()
        folder2 = self.folder2_var.get().strip()
        template = self.template_var.get().strip()
        output = self.output_var.get().strip()

        if not all([folder1, folder2, template, output]):
            messagebox.showerror("参数不完整", "请先完整选择全部输入项。")
            return

        self.start_btn.configure(state=tk.DISABLED)
        self.log("=" * 60)
        self.log("开始执行任务...")

        thread = threading.Thread(
            target=self._run_processor,
            args=(folder1, folder2, template, output),
            daemon=True,
        )
        thread.start()

    def _run_processor(self, folder1: str, folder2: str, template: str, output: str):
        def safe_log(msg: str):
            self.root.after(0, lambda: self.log(msg))

        try:
            processor = ExcelFlowProcessor(logger=safe_log)
            stats = processor.process(folder1, folder2, template, output)
            self.root.after(0, lambda: self._on_success(stats, output))
        except ProcessingError as exc:
            self.root.after(0, lambda: self._on_failure(str(exc)))
        except Exception as exc:  # noqa: BLE001
            self.root.after(0, lambda: self._on_failure(f"未预期异常: {exc}"))

    def _on_success(self, stats: ProcessStats, output: str):
        self.log("任务执行成功。")
        self.log(
            "结果汇总："
            f"文件1={stats.folder1_files}个, 文件2={stats.folder2_files}个, "
            f"现券买卖={stats.rows_spot_bond}行, "
            f"质押式回购={stats.rows_pledged_repo}行, "
            f"协议回购={stats.rows_agreement_repo}行"
        )
        self.start_btn.configure(state=tk.NORMAL)
        messagebox.showinfo("处理成功", f"数据处理完成，输出文件：\n{Path(output).resolve()}")

    def _on_failure(self, error_message: str):
        self.log(f"任务失败：{error_message}")
        self.start_btn.configure(state=tk.NORMAL)
        messagebox.showerror("处理失败", error_message)
