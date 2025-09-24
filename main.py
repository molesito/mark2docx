import io
import re
import json
import base64
import pathlib
from typing import Dict, List, Optional, Tuple

from flask import Flask, request, send_file, jsonify
from werkzeug.datastructures import FileStorage

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_SECTION
from docxcompose.composer import Composer
from PIL import Image

# Markdown parser
from markdown_it import MarkdownIt
from mdit_py_plugins import tasklists, table, anchors, sub, sup

# LaTeX -> MathML
from latex2mathml.converter import convert as latex_to_mathml
from lxml import etree

# Fallback render to image
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

app = Flask(__name__)

# -----------------------
# Markdown configuration
# -----------------------
md = (
    MarkdownIt("commonmark", {"html": False})
    .use(table.plugin)
    .use(tasklists.tasklists_plugin, enabled=True, label=True)
    .use(anchors.plugin)
    .use(sub.plugin)
    .use(sup.plugin)
)

MATH_INLINE_SPLIT_RE = re.compile(r"(\$\$.*?\$\$|\$.*?\$)", re.DOTALL)
NEW_PAGE_MARKER = "[NUEVA PÁGINA]"

# -----------------------
# MathML -> OMML with XSLT
# -----------------------
XSLT_PATH = pathlib.Path(__file__).parent / "MML2OMML.XSL"

def mathml_to_omml(mathml: str) -> str:
    xslt_root = etree.parse(str(XSLT_PATH))
    transform = etree.XSLT(xslt_root)
    mathml_tree = etree.fromstring(mathml.encode("utf-8"))
    omml_tree = transform(mathml_tree)
    return etree.tostring(omml_tree)

def append_omml(paragraph, omml_xml: str, as_block: bool):
    root = etree.fromstring(omml_xml.encode("utf-8"))
    paragraph._p.append(root)

def latex_to_omml_or_image(doc: Document, paragraph, latex_src: str, as_block: bool):
    try:
        if latex_src.startswith("$$") and latex_src.endswith("$$"):
            core = latex_src[2:-2].strip()
        elif latex_src.startswith("$") and latex_src.endswith("$"):
            core = latex_src[1:-1].strip()
        else:
            core = latex_src.strip()

        mathml = latex_to_mathml(core)
        omml_xml = mathml_to_omml(mathml).decode("utf-8")
        append_omml(paragraph, omml_xml, as_block=as_block)
        return
    except Exception:
        pass

    # fallback: render as image
    try:
        fig = plt.figure()
        ax = fig.add_subplot(111)
        ax.axis("off")
        render_tex = core
        if not (render_tex.startswith("$") and render_tex.endswith("$")):
            render_tex = f"${render_tex}$"
        ax.text(0.5, 0.5, render_tex, ha="center", va="center")
        buf = io.BytesIO()
        plt.savefig(buf, format="png", bbox_inches="tight", pad_inches=0.2, dpi=300)
        plt.close(fig)
        buf.seek(0)
        run = paragraph.add_run()
        run.add_picture(buf, width=Inches(5.5))
    except Exception:
        paragraph.add_run(core)

# -----------------------
# Helpers
# -----------------------
def ensure_code_style(doc: Document, style_name="Code"):
    if style_name in [s.name for s in doc.styles]:
        return style_name
    s = doc.styles.add_style(style_name, 1)
    s.font.name = "Consolas"
    s.font.size = Pt(10)
    return style_name

def set_default_font_black(run):
    run.font.color.rgb = RGBColor(0, 0, 0)

def add_page_break(doc: Document):
    p = doc.add_paragraph()
    r = p.add_run()
    r.add_break(type=1)  # WD_BREAK.PAGE

# -----------------------
# Inline renderer
# -----------------------
def _render_inline(doc: Document, paragraph, inline_token, images_map: Dict[str, str]):
    text = inline_token.content or ""
    if not text:
        return
    parts = MATH_INLINE_SPLIT_RE.split(text)
    for part in parts:
        if not part:
            continue
        if part.startswith("$$") and part.endswith("$$"):
            latex_to_omml_or_image(doc, paragraph, part, as_block=True)
            continue
        if part.startswith("$") and part.endswith("$") and len(part) >= 2:
            latex_to_omml_or_image(doc, paragraph, part, as_block=False)
            continue
        paragraph.add_run(part)

# -----------------------
# Core rendering
# -----------------------
def render_tokens_to_docx(doc: Document, tokens, env, images_map: Dict[str, str]):
    code_style = ensure_code_style(doc)
    i = 0
    while i < len(tokens):
        t = tokens[i]
        if t.type == "heading_open":
            level = int(t.tag[-1]) if t.tag.startswith("h") else 1
            p = doc.add_paragraph()
            p.style = f"Heading {min(level,5)}"
            i += 1
            if i < len(tokens) and tokens[i].type == "inline":
                _render_inline(doc, p, tokens[i], images_map)
                i += 1
            i += 1
            continue
        if t.type == "paragraph_open":
            if i+1 < len(tokens) and tokens[i+1].type == "inline":
                page_text = tokens[i+1].content.strip()
                if NEW_PAGE_MARKER in page_text:
                    add_page_break(doc)
                    i += 3
                    continue
            p = doc.add_paragraph()
            i += 1
            if i < len(tokens) and tokens[i].type == "inline":
                _render_inline(doc, p, tokens[i], images_map)
                i += 1
            i += 1
            continue
        if t.type == "fence" or t.type == "code_block":
            code = t.content
            p = doc.add_paragraph()
            p.style = code_style
            for line in code.splitlines():
                r = p.add_run(line)
                r.font.name = "Consolas"
                set_default_font_black(r)
                p = doc.add_paragraph()
                p.style = code_style
            i += 1
            continue
        i += 1

# -----------------------
# /docx endpoint
# -----------------------
@app.route("/docx", methods=["POST"])
def make_docx():
    images_map: Dict[str, str] = {}
    markdown_text = ""
    if request.content_type and "application/json" in request.content_type:
        data = request.get_json(force=True, silent=True) or {}
        markdown_text = data.get("markdown", "") or ""
        images_map = data.get("images", {}) or {}
        output_name = data.get("output_name", "documento.docx")
    else:
        markdown_text = request.form.get("markdown", "") or ""
        output_name = request.form.get("output_name", "documento.docx")
    if not markdown_text.strip():
        return jsonify({"error": "Falta 'markdown'."}), 400
    doc = Document()
    env = {}
    tokens = md.parse(markdown_text, env)
    render_tokens_to_docx(doc, tokens, env, images_map)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=output_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

# -----------------------
# /merge endpoint
# -----------------------
@app.route("/merge", methods=["POST"])
def merge_docx():
    output_name = "merged.docx"
    streams: List[io.BytesIO] = []
    if request.content_type and "application/json" in request.content_type:
        data = request.get_json(force=True, silent=True) or {}
        output_name = data.get("output_name", "merged.docx")
        files_b64 = data.get("files", []) or []
        for b in files_b64:
            try:
                streams.append(io.BytesIO(base64.b64decode(b)))
            except Exception:
                return jsonify({"error": "Archivo base64 inválido."}), 400
    else:
        if "files" not in request.files:
            return jsonify({"error": "Envía ficheros en el campo 'files'."}), 400
        fs: List[FileStorage] = request.files.getlist("files")
        for f in fs:
            streams.append(io.BytesIO(f.read()))
        output_name = request.form.get("output_name", "merged.docx") or "merged.docx"
    if not streams:
        return jsonify({"error": "No se han recibido documentos."}), 400
    base_doc = Document(streams[0])
    composer = Composer(base_doc)
    for st in streams[1:]:
        subdoc = Document(st)
        base_doc.add_section(WD_SECTION.NEW_PAGE)
        composer.append(subdoc)
    out = io.BytesIO()
    composer.save(out)
    out.seek(0)
    return send_file(
        out,
        as_attachment=True,
        download_name=output_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

@app.route("/", methods=["GET"])
def health():
    return jsonify({"ok": True})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
