"""Vercel serverless function: convert Markdown text to DOCX."""

from http.server import BaseHTTPRequestHandler
import json
import io

import markdown
from docx import Document
from docx.shared import Pt
from htmldocx import HtmlToDocx


def convert(md_text):
    """Convert markdown text to docx bytes."""
    html = markdown.markdown(
        md_text,
        extensions=["tables", "fenced_code", "codehilite", "toc", "sane_lists"],
    )

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    parser = HtmlToDocx()
    parser.add_html_to_document(html, doc)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length))

            md_text = body.get("content", "")
            filename = body.get("filename", "document.md")

            if not md_text.strip():
                self.send_response(400)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": "Empty markdown content"}).encode())
                return

            docx_bytes = convert(md_text)
            docx_name = filename.rsplit(".", 1)[0] + ".docx"

            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{docx_name}"')
            self.send_header("Content-Length", str(len(docx_bytes)))
            self.end_headers()
            self.wfile.write(docx_bytes)

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode())

    def log_message(self, format, *args):
        pass  # suppress default stderr logging
