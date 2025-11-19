"""简单的可视化界面，方便非技术用户批量生成 Word 报告。"""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from generate_word_report import generate_word_report


class ReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("图片生成 Word 报告")
        self.geometry("520x220")
        self.resizable(False, False)

        self.folder_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self._folder_path: Path | None = None
        self._output_path: Path | None = None

        self._build_form()

    def _build_form(self) -> None:
        padding = {"padx": 10, "pady": 8}

        tk.Label(self, text="图片文件夹：").grid(row=0, column=0, sticky="e", **padding)
        tk.Entry(self, textvariable=self.folder_var, width=45).grid(
            row=0, column=1, **padding
        )
        tk.Button(self, text="浏览...", command=self.select_folder).grid(
            row=0, column=2, **padding
        )

        tk.Label(self, text="输出 Word：").grid(row=1, column=0, sticky="e", **padding)
        tk.Entry(self, textvariable=self.output_var, width=45).grid(
            row=1, column=1, **padding
        )
        tk.Button(self, text="选择...", command=self.select_output).grid(
            row=1, column=2, **padding
        )

        self.status_var = tk.StringVar(value="请选择图片文件夹并设置输出路径")
        tk.Label(self, textvariable=self.status_var, fg="#555").grid(
            row=2, column=0, columnspan=3, sticky="w", padx=10, pady=5
        )

        self.progress = ttk.Progressbar(
            self, length=460, mode="determinate", maximum=100, value=0
        )
        self.progress.grid(row=3, column=0, columnspan=3, padx=10, pady=5)

        self.run_button = tk.Button(self, text="开始生成", command=self.run_generation)
        self.run_button.grid(row=4, column=0, columnspan=3, pady=15)

    def select_folder(self) -> None:
        path = filedialog.askdirectory(title="选择包含图片的文件夹")
        if path:
            self._folder_path = Path(path)
            self.folder_var.set(self._folder_path.as_posix())
            if not self.output_var.get():
                self._output_path = self._folder_path / "图片汇总.docx"
                self.output_var.set(self._output_path.as_posix())

    def select_output(self) -> None:
        initial_dir = (
            self._folder_path
            if self._folder_path is not None
            else (Path(self.folder_var.get()) if self.folder_var.get() else Path.cwd())
        )
        path = filedialog.asksaveasfilename(
            title="选择输出 Word",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialdir=initial_dir,
            initialfile="图片汇总.docx",
        )
        if path:
            self._output_path = Path(path)
            self.output_var.set(self._output_path.as_posix())

    @staticmethod
    def _parse_path(text: str, fallback: Path | None = None) -> Path | None:
        value = text.strip()
        if not value:
            return fallback
        if len(value) >= 2 and value[0].isalpha() and value[1] in ("/", "\\"):
            if len(value) == 2 or value[2] not in (":", "/", "\\"):
                value = f"{value[0]}:{value[1:]}"
        normalized = value.replace("\\", "/")
        try:
            return Path(normalized)
        except Exception:
            return fallback

    def run_generation(self) -> None:
        folder = self._parse_path(self.folder_var.get(), self._folder_path)
        output_path = self._parse_path(self.output_var.get(), self._output_path)

        if folder is not None:
            self._folder_path = folder
        if output_path is not None:
            self._output_path = output_path

        if folder is None or not folder.exists():
            messagebox.showerror("提示", "请选择有效的图片文件夹")
            return

        if output_path is None:
            output_path = folder / "图片汇总.docx"
        else:
            output_path.parent.mkdir(parents=True, exist_ok=True)

        self.run_button.config(state="disabled")
        self.status_var.set("正在生成 Word，请稍候…")
        self.progress.config(value=0, maximum=100)

        def progress_callback(done: int, total: int) -> None:
            self.after(0, lambda: self._update_progress(done, total))

        def task() -> None:
            try:
                generate_word_report(
                    folder, output_path, progress_callback=progress_callback
                )
            except Exception as exc:  # noqa: BLE001
                error_msg = f"生成失败：{exc}\n请检查图片与输出路径后重试。"
                self.after(0, lambda msg=error_msg: self._on_finish(False, msg))
            else:
                self.after(
                    0,
                    lambda: self._on_finish(
                        True, f"生成成功！已保存到：\n{output_path}"
                    ),
                )

        threading.Thread(target=task, daemon=True).start()

    def _on_finish(self, success: bool, message: str) -> None:
        self.run_button.config(state="normal")
        self.progress.stop()
        self.status_var.set(message)
        if success:
            self.progress.config(value=self.progress["maximum"])
        if success:
            messagebox.showinfo("完成", message)
        else:
            messagebox.showerror("错误", message)

    def _update_progress(self, done: int, total: int) -> None:
        if total <= 0:
            self.progress.config(mode="indeterminate")
            self.progress.start(10)
            return

        self.progress.stop()
        self.progress.config(mode="determinate", maximum=total, value=done)
        percent = done / total * 100 if total else 0
        self.status_var.set(f"正在生成：{done}/{total} ({percent:4.1f}%)")


def main() -> None:
    app = ReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()

