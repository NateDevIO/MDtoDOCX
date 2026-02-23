# MD to DOCX Converter

A simple Python app that converts Markdown (`.md`) files to Word (`.docx`) format.

## Requirements

- Python 3
- Install dependencies: `pip install python-docx markdown htmldocx`

## Usage

**GUI mode** — run without arguments to open a file picker:

```
python md2docx.py
```

**Command-line mode** — pass files or folders as arguments:

```
python md2docx.py myfile.md
python md2docx.py ./docs/
```

The `.docx` files are saved alongside the originals.

## Supported Formatting

Headings, bold, italic, bullet lists, code blocks, tables, blockquotes, and links.
