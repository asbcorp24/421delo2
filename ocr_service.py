"""Minimal HTTP wrapper around OnnxTR for the document-management application.

Run this file inside the official OnnxTR Docker image. It intentionally uses
only Python's standard library because the base image is an OCR library, not a
ready-made REST service.
"""
import cgi
import json
import os
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from onnxtr.io import DocumentFile
from onnxtr.models import from_hub, ocr_predictor


MODEL = None
MODEL_LOCK = threading.Lock()


def get_model():
    global MODEL
    with MODEL_LOCK:
        if MODEL is None:
            # Multilingual PARSeq is necessary for Russian business documents.
            recognition_model = from_hub("Felix92/onnxtr-parseq-multilingual-v1")
            MODEL = ocr_predictor(det_arch="fast_base", reco_arch=recognition_model)
    return MODEL


def recognize(path):
    suffix = Path(path).suffix.lower()
    if suffix == ".pdf":
        document = DocumentFile.from_pdf(path)
    else:
        document = DocumentFile.from_images(path)
    # Inference is serialised to avoid simultaneous model initialisation and
    # excessive memory consumption on a CPU-only host.
    model = get_model()
    with MODEL_LOCK:
        result = model(document)
    exported = result.export_as("text")
    return exported if isinstance(exported, str) else str(exported)


class OcrHandler(BaseHTTPRequestHandler):
    def send_json(self, status, payload):
        encoded = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def do_GET(self):
        if self.path.rstrip("/") == "/health":
            self.send_json(200, {"status": "ok"})
        else:
            self.send_json(404, {"error": "not found"})

    def do_POST(self):
        if self.path.rstrip("/") != "/ocr":
            self.send_json(404, {"error": "not found"})
            return
        if "multipart/form-data" not in (self.headers.get("Content-Type") or ""):
            self.send_json(400, {"error": "multipart/form-data expected"})
            return

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers["Content-Type"]},
        )
        uploaded = form["file"] if "file" in form else None
        if uploaded is None or not getattr(uploaded, "file", None) or not uploaded.filename:
            self.send_json(400, {"error": "file is required"})
            return

        suffix = Path(uploaded.filename).suffix.lower()
        if suffix not in {".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}:
            self.send_json(415, {"error": "only PDF and image files are supported"})
            return

        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
                temp_file.write(uploaded.file.read())
                temp_path = temp_file.name
            self.send_json(200, {"text": recognize(temp_path)})
        except Exception as exc:
            self.send_json(500, {"error": str(exc)})
        finally:
            if temp_path:
                try:
                    os.unlink(temp_path)
                except OSError:
                    pass

    def log_message(self, format_string, *args):
        print("%s - %s" % (self.address_string(), format_string % args))


if __name__ == "__main__":
    port = int(os.getenv("OCR_PORT", "8200"))
    print(f"OnnxTR OCR service is listening on port {port}")
    ThreadingHTTPServer(("0.0.0.0", port), OcrHandler).serve_forever()
