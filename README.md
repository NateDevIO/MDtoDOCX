# MD to DOCX Converter

A simple Markdown to Word (.docx) converter — available as a desktop app, CLI tool, or hosted web app.

**Live demo:** [mdtodocx.vercel.app](https://mdtodocx.vercel.app)

## Web App

Upload a `.md` file in the browser and download the converted `.docx`. Hosted on Vercel with a Python serverless backend.

## Desktop / CLI

### Requirements

- Python 3
- `pip install python-docx markdown htmldocx`

### Usage

**GUI mode** — run without arguments to open a file picker:

```
python md2docx.py
```

**CLI mode** — pass files or folders as arguments:

```
python md2docx.py myfile.md
python md2docx.py ./docs/
```

Output `.docx` files are saved alongside the originals.

## Supported Formatting

Headings, bold, italic, bullet lists, code blocks, tables, blockquotes, and links.
