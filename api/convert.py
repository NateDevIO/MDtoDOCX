"""Vercel serverless function: convert Markdown text to DOCX."""

from http.server import BaseHTTPRequestHandler
import json
import io
import traceback


def convert(md_text):
    """Convert markdown text to docx bytes."""
    import markdown
    from docx import Document
    from docx.shared import Pt
    from htmldocx import HtmlToDocx

    html = markdown.markdown(
        md_text,
        extensions=["tables", "fenced_code", "toc", "sane_lists"],
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
    def do_GET(self):
        """Health check — also tests that imports work."""
        try:
            import markdown
            from docx import Document
            from htmldocx import HtmlToDocx

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"status": "ok"}).encode())
        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({
                "status": "error",
                "error": str(e),
                "trace": traceback.format_exc(),
            }).encode())

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length))

            md_text = body.get("content", "")
            filename = body.get("filename", "document.md")

            if not md_text.strip():
                self._json_error(400, "Empty markdown content")
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
            self._json_error(500, str(e), traceback.format_exc())

    def _json_error(self, code, message, trace=None):
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        payload = {"error": message}
        if trace:
            payload["trace"] = trace
        self.wfile.write(json.dumps(payload).encode())

    def log_message(self, format, *args):
        pass
