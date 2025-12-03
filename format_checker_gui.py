"""Простое десктоп-приложение на Tkinter для проверки оформления DOCX."""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from typing import List, Tuple

import docx

from format_checker_core import Issue, annotate_document, check_document


class FormatCheckerGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Проверка оформления статей (DOCX)")
        self.geometry("1100x700")

        self.selected_file: str | None = None
        self.issues: List[Issue] = []
        self.paragraph_ranges: List[Tuple[str, str]] = []
        self.annotated_path: str | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        top_frame = ttk.Frame(self)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        open_btn = ttk.Button(top_frame, text="Открыть файл", command=self.open_file)
        open_btn.pack(side=tk.LEFT)

        self.file_var = tk.StringVar()
        file_entry = ttk.Entry(top_frame, textvariable=self.file_var, width=80)
        file_entry.state(["readonly"])
        file_entry.pack(side=tk.LEFT, padx=5)

        self.check_btn = ttk.Button(top_frame, text="Проверить", command=self.run_check)
        self.check_btn.pack(side=tk.LEFT)

        self.open_annotated_btn = ttk.Button(top_frame, text="Открыть аннотированный DOCX", command=self.open_annotated, state=tk.DISABLED)
        self.open_annotated_btn.pack(side=tk.LEFT, padx=5)

        center_frame = ttk.Frame(self)
        center_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Левая панель: список замечаний
        left_frame = ttk.Frame(center_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        ttk.Label(left_frame, text="Замечания:").pack(anchor=tk.W)
        self.issue_list = tk.Listbox(left_frame, width=55)
        self.issue_list.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        self.issue_list.bind("<<ListboxSelect>>", self.on_issue_select)

        issue_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.issue_list.yview)
        issue_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.issue_list.config(yscrollcommand=issue_scroll.set)

        # Правая панель: текст документа
        right_frame = ttk.Frame(center_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)

        ttk.Label(right_frame, text="Текст документа:").pack(anchor=tk.W)
        self.text_widget = ScrolledText(right_frame, wrap=tk.WORD)
        self.text_widget.pack(fill=tk.BOTH, expand=True)
        self.text_widget.config(state=tk.DISABLED)

        self.text_widget.tag_configure("error_paragraph", background="#f8d7da")
        self.text_widget.tag_configure("warn_paragraph", background="#fff3cd")
        self.text_widget.tag_configure("selected_issue", background="#cce5ff")

        # Строка статуса
        self.status_var = tk.StringVar(value="Файл не выбран")
        status_bar = ttk.Label(self, textvariable=self.status_var, anchor=tk.W)
        status_bar.pack(fill=tk.X, padx=10, pady=5)

    def open_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("DOCX files", "*.docx")])
        if not path:
            return
        self.selected_file = path
        self.file_var.set(path)
        self.status_var.set(f"Файл: {os.path.basename(path)}")
        self.annotated_path = None
        self.open_annotated_btn.config(state=tk.DISABLED)

    def run_check(self) -> None:
        if not self.selected_file:
            messagebox.showerror("Ошибка", "Сначала выберите DOCX-файл")
            return

        self.check_btn.config(state=tk.DISABLED)
        self.update_idletasks()
        try:
            self.issues = check_document(self.selected_file)
            self.populate_issue_list()
            self.load_document_preview()
            self.update_status()
            if any(issue.level != "OK" for issue in self.issues):
                self.annotated_path = annotate_document(self.selected_file, self.issues)
                self.open_annotated_btn.config(state=tk.NORMAL)
            else:
                self.annotated_path = None
                self.open_annotated_btn.config(state=tk.DISABLED)
        except Exception as exc:  # pragma: no cover - защита от падений в GUI
            messagebox.showerror("Ошибка", f"Не удалось выполнить проверку: {exc}")
        finally:
            self.check_btn.config(state=tk.NORMAL)

    def populate_issue_list(self) -> None:
        self.issue_list.delete(0, tk.END)
        self.issue_items: List[Issue] = []
        for issue in self.issues:
            if issue.level == "OK":
                continue
            para_part = f" (абз. {issue.paragraph_index + 1})" if issue.paragraph_index is not None else ""
            text = f"[{issue.level}] {issue.rule}{para_part} {issue.message}"
            self.issue_list.insert(tk.END, text)
            self.issue_items.append(issue)

    def load_document_preview(self) -> None:
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.delete("1.0", tk.END)
        self.paragraph_ranges = []

        doc = docx.Document(self.selected_file)
        for idx, paragraph in enumerate(doc.paragraphs):
            start_index = self.text_widget.index(tk.END)
            display_text = paragraph.text or ""
            self.text_widget.insert(tk.END, display_text + "\n")
            end_index = self.text_widget.index(tk.END)
            self.paragraph_ranges.append((start_index, end_index))

        self.apply_highlighting()
        self.text_widget.config(state=tk.DISABLED)

    def apply_highlighting(self) -> None:
        self.text_widget.tag_remove("error_paragraph", "1.0", tk.END)
        self.text_widget.tag_remove("warn_paragraph", "1.0", tk.END)
        self.text_widget.tag_remove("selected_issue", "1.0", tk.END)

        error_paragraphs = {issue.paragraph_index for issue in self.issues if issue.level == "ERROR" and issue.paragraph_index is not None}
        warn_paragraphs = {issue.paragraph_index for issue in self.issues if issue.level == "WARN" and issue.paragraph_index is not None}

        for idx in error_paragraphs:
            if idx is None or idx >= len(self.paragraph_ranges):
                continue
            start, end = self.paragraph_ranges[idx]
            self.text_widget.tag_add("error_paragraph", start, end)

        for idx in warn_paragraphs:
            if idx is None or idx in error_paragraphs or idx >= len(self.paragraph_ranges):
                continue
            start, end = self.paragraph_ranges[idx]
            self.text_widget.tag_add("warn_paragraph", start, end)

    def on_issue_select(self, event: tk.Event) -> None:
        if not getattr(self, "issue_items", None):
            return
        selection = self.issue_list.curselection()
        if not selection:
            return
        issue = self.issue_items[selection[0]]
        if issue.paragraph_index is None or issue.paragraph_index >= len(self.paragraph_ranges):
            return

        start, end = self.paragraph_ranges[issue.paragraph_index]
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.tag_remove("selected_issue", "1.0", tk.END)
        self.text_widget.tag_add("selected_issue", start, end)
        self.text_widget.see(start)
        self.text_widget.config(state=tk.DISABLED)

    def update_status(self) -> None:
        totals = {"ERROR": 0, "WARN": 0, "OK": 0}
        for issue in self.issues:
            totals[issue.level] = totals.get(issue.level, 0) + 1
        file_part = os.path.basename(self.selected_file) if self.selected_file else "Не выбран"
        self.status_var.set(
            f"Файл: {file_part} | Ошибки: {totals.get('ERROR', 0)} | Предупреждения: {totals.get('WARN', 0)} | OK: {totals.get('OK', 0)}"
        )

    def open_annotated(self) -> None:
        if not self.annotated_path:
            return
        try:
            os.startfile(self.annotated_path)
        except OSError:
            messagebox.showerror("Ошибка", "Не удалось открыть аннотированный файл")


def main() -> None:
    app = FormatCheckerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
