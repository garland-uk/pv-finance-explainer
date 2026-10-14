"""Builds ../index.html, a single self-contained page, from the files in this folder.

Run from anywhere:  python3 source/build.py
Edit styles.css, app.js or template.html, then run this again and commit index.html.
"""
import base64
import pathlib

HERE = pathlib.Path(__file__).resolve().parent
OUT = HERE.parent / "index.html"


def b64(path):
    return base64.b64encode((HERE / path).read_bytes()).decode("ascii")


def data_uri(path, mime):
    return f"data:{mime};base64,{b64(path)}"


def safe_inline(js, name):
    if "</script" in js.lower() or "<!--" in js:
        raise SystemExit(f"{name} contains a sequence that cannot be inlined safely")
    return js


css = (HERE / "styles.css").read_text(encoding="utf-8")
css = css.replace("__ROBOTO900__", data_uri("vendor/roboto-latin-900-normal.woff2", "font/woff2"))

html = (HERE / "template.html").read_text(encoding="utf-8")
parts = {
    "__CSS__": css,
    "__LOGO_H_DATAURI__": data_uri("assets/garland-logo-horizontal.png", "image/png"),
    "__LOGO_REV_DATAURI__": data_uri("assets/garland-logo-reversed.png", "image/png"),
    "__LOGO_SQ_DATAURI__": data_uri("assets/garland-logo-square.png", "image/png"),
    "__HERO_DATAURI__": data_uri("assets/hero-solar.jpg", "image/jpeg"),
    "__PDFJS_B64__": b64("vendor/pdf.min.mjs"),
    "__PDFWORKER_B64__": b64("vendor/pdf.worker.min.mjs"),
    "__JSZIP__": safe_inline((HERE / "vendor/jszip.min.js").read_text(encoding="utf-8"), "jszip"),
    "__APP__": safe_inline((HERE / "app.js").read_text(encoding="utf-8"), "app.js"),
}
for key, value in parts.items():
    html = html.replace(key, value)

OUT.write_text(html, encoding="utf-8")
print(f"Built {OUT} ({OUT.stat().st_size / 1e6:.1f} MB)")
