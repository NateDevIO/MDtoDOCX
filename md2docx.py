"""
MD to DOCX Converter
Converts Markdown files to Word (.docx) format.
"""

import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

import markdown
from docx import Document
from docx.shared import Pt, Inches
from htmldocx import HtmlToDocx


def convert_md_to_docx(md_path, docx_path=None):
    """Convert a Markdown file to a .docx file."""
    md_path = Path(md_path)
    if docx_path is None:
        docx_path = md_path.with_suffix(".docx")
    else:
        docx_path = Path(docx_path)

    md_text = md_path.read_text(encoding="utf-8")

    # Convert markdown to HTML with common extensions
    html = markdown.markdown(
        md_text,
        extensions=["tables", "fenced_code", "codehilite", "toc", "sane_lists"],
    )

    doc = Document()

    # Set default font
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)

    # Convert HTML into the docx document
    parser = HtmlToDocx()
    parser.add_html_to_document(html, doc)

    doc.save(str(docx_path))
    return docx_path


class App:
    def __init__(self, root):
        self.root = root
        root.title("MD to DOCX Converter")
        root.geometry("520x320")
        root.resizable(False, False)

        frame = tk.Frame(root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Markdown to Word Converter", font=("Segoe UI", 14, "bold")).pack(pady=(0, 12))
        tk.Label(frame, text="Select one or more .md files to convert to .docx", font=("Segoe UI", 10)).pack()

        btn_frame = tk.Frame(frame)
        btn_frame.pack(pady=18)

        tk.Button(btn_frame, text="Select Files...", command=self.pick_files, width=16, font=("Segoe UI", 10)).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Select Folder...", command=self.pick_folder, width=16, font=("Segoe UI", 10)).pack(side="left", padx=6)

        self.status = tk.Text(frame, height=8, width=58, state="disabled", font=("Consolas", 9))
        self.status.pack(pady=(4, 0))

    def log(self, msg):
        self.status.config(state="normal")
        self.status.insert("end", msg + "\n")
        self.status.see("end")
        self.status.config(state="disabled")

    def convert(self, paths):
        converted = 0
        for p in paths:
            try:
                out = convert_md_to_docx(p)
                self.log(f"OK:  {Path(p).name}  ->  {out.name}")
                converted += 1
            except Exception as e:
                self.log(f"ERR: {Path(p).name}  ({e})")
        self.log(f"--- Done: {converted}/{len(paths)} converted ---")

    def pick_files(self):
        files = filedialog.askopenfilenames(
            title="Select Markdown files",
            filetypes=[("Markdown files", "*.md"), ("All files", "*.*")],
        )
        if files:
            self.convert(files)

    def pick_folder(self):
        folder = filedialog.askdirectory(title="Select folder with .md files")
        if folder:
            md_files = sorted(Path(folder).glob("*.md"))
            if not md_files:
                messagebox.showinfo("No files", "No .md files found in that folder.")
                return
            self.convert(md_files)


def main():
    # CLI mode: pass file paths as arguments
    if len(sys.argv) > 1:
        for arg in sys.argv[1:]:
            p = Path(arg)
            if p.is_file() and p.suffix.lower() == ".md":
                out = convert_md_to_docx(p)
                print(f"Converted: {p} -> {out}")
            elif p.is_dir():
                for md in sorted(p.glob("*.md")):
                    out = convert_md_to_docx(md)
                    print(f"Converted: {md} -> {out}")
            else:
                print(f"Skipped: {arg}")
        return

    # GUI mode
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
