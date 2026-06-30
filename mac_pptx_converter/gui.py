"""
macOS GUI for the Mac PPTX Converter.

Provides a drag-and-drop style window for converting
Windows VBA macro-enabled Office files for Mac compatibility.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, font as tkfont
from pathlib import Path

from .converter import convert_file, format_report


SUPPORTED_EXTENSIONS = (".pptm", ".pptx", ".xlsm", ".xlsx", ".docm", ".docx")

BG_COLOR = "#1e1e2e"
FG_COLOR = "#cdd6f4"
ACCENT_COLOR = "#89b4fa"
SUCCESS_COLOR = "#a6e3a1"
WARNING_COLOR = "#f9e2af"
ERROR_COLOR = "#f38ba8"
DROP_ZONE_BG = "#313244"
DROP_ZONE_BORDER = "#585b70"
BUTTON_BG = "#89b4fa"
BUTTON_FG = "#1e1e2e"


class ConverterApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Mac PPTX Converter")
        self.root.configure(bg=BG_COLOR)
        self.root.minsize(640, 520)
        self.root.geometry("720x600")

        if sys.platform == "darwin":
            self.root.createcommand("tk::mac::OpenDocument", self._on_open_document)

        self._setup_fonts()
        self._build_ui()
        self._setup_dnd()

    def _setup_fonts(self):
        families = tkfont.families()
        if "SF Pro Text" in families:
            base = "SF Pro Text"
        elif "Helvetica Neue" in families:
            base = "Helvetica Neue"
        else:
            base = "Helvetica"

        mono_candidates = ["SF Mono", "Menlo", "Monaco", "Courier"]
        mono = next((f for f in mono_candidates if f in families), "Courier")

        self.font_normal = (base, 13)
        self.font_bold = (base, 13, "bold")
        self.font_title = (base, 20, "bold")
        self.font_subtitle = (base, 11)
        self.font_mono = (mono, 11)

    def _build_ui(self):
        main = tk.Frame(self.root, bg=BG_COLOR, padx=24, pady=16)
        main.pack(fill=tk.BOTH, expand=True)

        title = tk.Label(
            main, text="Mac PPTX Converter", font=self.font_title,
            fg=FG_COLOR, bg=BG_COLOR,
        )
        title.pack(pady=(0, 2))

        subtitle = tk.Label(
            main,
            text="Windows VBA macros to Mac compatible format",
            font=self.font_subtitle, fg=DROP_ZONE_BORDER, bg=BG_COLOR,
        )
        subtitle.pack(pady=(0, 16))

        self.drop_frame = tk.Frame(
            main, bg=DROP_ZONE_BG, highlightbackground=DROP_ZONE_BORDER,
            highlightthickness=2, cursor="hand2",
        )
        self.drop_frame.pack(fill=tk.X, ipady=32)

        self.drop_label = tk.Label(
            self.drop_frame,
            text="Click to select files\n\n.pptm  .xlsm  .docm  .pptx  .xlsx  .docx",
            font=self.font_normal, fg=ACCENT_COLOR, bg=DROP_ZONE_BG,
            justify=tk.CENTER,
        )
        self.drop_label.pack(expand=True, fill=tk.BOTH, padx=16, pady=16)

        for widget in (self.drop_frame, self.drop_label):
            widget.bind("<Button-1>", self._on_click_select)

        btn_frame = tk.Frame(main, bg=BG_COLOR)
        btn_frame.pack(fill=tk.X, pady=(12, 8))

        self.convert_btn = tk.Button(
            btn_frame, text="Convert Selected Files", font=self.font_bold,
            bg=BUTTON_BG, fg=BUTTON_FG, activebackground=ACCENT_COLOR,
            activeforeground=BUTTON_FG, relief=tk.FLAT, padx=20, pady=8,
            command=self._on_convert, state=tk.DISABLED,
        )
        self.convert_btn.pack(side=tk.LEFT)

        self.status_label = tk.Label(
            btn_frame, text="", font=self.font_subtitle,
            fg=FG_COLOR, bg=BG_COLOR, anchor=tk.W,
        )
        self.status_label.pack(side=tk.LEFT, padx=(16, 0), fill=tk.X, expand=True)

        self.report_text = scrolledtext.ScrolledText(
            main, font=self.font_mono, bg="#181825", fg=FG_COLOR,
            insertbackground=FG_COLOR, relief=tk.FLAT,
            highlightthickness=1, highlightbackground=DROP_ZONE_BORDER,
            wrap=tk.WORD, height=14,
        )
        self.report_text.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        self.report_text.tag_configure("fixed", foreground=SUCCESS_COLOR)
        self.report_text.tag_configure("warning", foreground=WARNING_COLOR)
        self.report_text.tag_configure("error", foreground=ERROR_COLOR)
        self.report_text.tag_configure("header", foreground=ACCENT_COLOR, font=(*self.font_mono[:1], self.font_mono[1], "bold"))
        self.report_text.configure(state=tk.DISABLED)

        self.selected_files: list[str] = []

    def _setup_dnd(self):
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
            if isinstance(self.root, TkinterDnD.Tk):
                self.drop_frame.drop_target_register(DND_FILES)
                self.drop_frame.dnd_bind("<<Drop>>", self._on_dnd_drop)
                self.drop_label.configure(
                    text="Drop files here or click to select\n\n"
                    ".pptm  .xlsm  .docm  .pptx  .xlsx  .docx"
                )
        except ImportError:
            pass

    def _on_open_document(self, *args):
        files = [f for f in args if os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS]
        if files:
            self.selected_files = files
            self._update_file_list()
            self._on_convert()

    def _on_dnd_drop(self, event):
        raw = event.data
        files = []
        if "{" in raw:
            import re
            files = re.findall(r"\{([^}]+)\}", raw)
            remaining = re.sub(r"\{[^}]+\}", "", raw).strip()
            if remaining:
                files.extend(remaining.split())
        else:
            files = raw.split()

        valid = [f for f in files if os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS]
        if valid:
            self.selected_files = valid
            self._update_file_list()

    def _on_click_select(self, event=None):
        files = filedialog.askopenfilenames(
            title="Select Office files to convert",
            filetypes=[
                ("Macro-enabled Office files", "*.pptm *.xlsm *.docm"),
                ("All Office files", "*.pptm *.pptx *.xlsm *.xlsx *.docm *.docx"),
                ("All files", "*.*"),
            ],
        )
        if files:
            self.selected_files = list(files)
            self._update_file_list()

    def _update_file_list(self):
        n = len(self.selected_files)
        names = [os.path.basename(f) for f in self.selected_files]
        if n == 1:
            self.drop_label.configure(text=f"Selected: {names[0]}")
        elif n <= 3:
            self.drop_label.configure(text="Selected:\n" + "\n".join(names))
        else:
            self.drop_label.configure(text=f"Selected: {n} files\n" + "\n".join(names[:3]) + "\n...")
        self.convert_btn.configure(state=tk.NORMAL)

    def _on_convert(self):
        if not self.selected_files:
            return
        self.convert_btn.configure(state=tk.DISABLED)
        self.status_label.configure(text="Converting...", fg=ACCENT_COLOR)
        self.report_text.configure(state=tk.NORMAL)
        self.report_text.delete("1.0", tk.END)
        self.report_text.configure(state=tk.DISABLED)

        threading.Thread(target=self._convert_thread, daemon=True).start()

    def _convert_thread(self):
        results = []
        for i, path in enumerate(self.selected_files):
            self.root.after(0, self._update_status,
                           f"Converting {i + 1}/{len(self.selected_files)}: {os.path.basename(path)}")
            base, ext = os.path.splitext(path)
            output_path = f"{base}_mac{ext}"
            result = convert_file(path, output_path)
            report = format_report(result)
            results.append((result, report))
            self.root.after(0, self._append_report, result, report)

        total_fixed = sum(r.fixed_count for r, _ in results)
        total_warnings = sum(r.warning_count for r, _ in results)
        total_errors = sum(r.error_count for r, _ in results)
        conv_errors = sum(1 for r, _ in results if r.errors)

        if conv_errors:
            summary = f"Done with errors. {total_fixed} fixed, {total_warnings} warnings, {total_errors} manual fixes needed"
            color = ERROR_COLOR
        elif total_errors:
            summary = f"Done. {total_fixed} auto-fixed, {total_errors} items need manual attention"
            color = WARNING_COLOR
        elif total_fixed:
            summary = f"Done. {total_fixed} items auto-fixed for Mac compatibility"
            color = SUCCESS_COLOR
        else:
            summary = "Done. No compatibility issues found"
            color = SUCCESS_COLOR

        self.root.after(0, self._finish, summary, color)

    def _update_status(self, text):
        self.status_label.configure(text=text, fg=ACCENT_COLOR)

    def _append_report(self, result, report_text):
        self.report_text.configure(state=tk.NORMAL)

        for line in report_text.split("\n"):
            stripped = line.strip()
            if stripped.startswith("=") or stripped.startswith("-"):
                self.report_text.insert(tk.END, line + "\n", "header")
            elif stripped.startswith("✓"):
                self.report_text.insert(tk.END, line + "\n", "fixed")
            elif stripped.startswith("⚠"):
                self.report_text.insert(tk.END, line + "\n", "warning")
            elif stripped.startswith("✗"):
                self.report_text.insert(tk.END, line + "\n", "error")
            elif "auto-fixed" in stripped or "Mac-compatible" in stripped:
                self.report_text.insert(tk.END, line + "\n", "fixed")
            else:
                self.report_text.insert(tk.END, line + "\n")

        self.report_text.configure(state=tk.DISABLED)
        self.report_text.see(tk.END)

    def _finish(self, summary, color):
        self.status_label.configure(text=summary, fg=color)
        self.convert_btn.configure(state=tk.NORMAL)

        if self.selected_files:
            first_output = os.path.splitext(self.selected_files[0])[0] + "_mac" + os.path.splitext(self.selected_files[0])[1]
            if os.path.exists(first_output) and sys.platform == "darwin":
                folder = os.path.dirname(first_output)
                os.system(f'open "{folder}"')

    def run(self):
        self.root.mainloop()


def main():
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
        app = ConverterApp.__new__(ConverterApp)
        app.root = root
        app.root.title("Mac PPTX Converter")
        app.root.configure(bg=BG_COLOR)
        app.root.minsize(640, 520)
        app.root.geometry("720x600")
        if sys.platform == "darwin":
            app.root.createcommand("tk::mac::OpenDocument", app._on_open_document)
        app._setup_fonts()
        app._build_ui()
        app._setup_dnd()
        app.run()
    except ImportError:
        app = ConverterApp()
        app.run()


if __name__ == "__main__":
    main()
