import io
import re
import base64
import json
from typing import Dict, List, Optional

from flask import Flask, request, send_file, jsonify
from werkzeug.datastructures import FileStorage

from docx import Document
from docx.shared import RGBColor, Inches, Pt
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docxcompose.composer import Composer

from markdown_it import MarkdownIt
from latex2mathml.converter import convert as latex2mathml
import mathml2omml
import lxml.etree as ET
from PIL import Image as PILImage


app = Flask(__name__)

# ------------------ Utilidades de estilo ------------------

def force_styles_black(doc: Document):
    """Asegura texto negro en estilos comunes de Word."""
    target_styles = ["Normal", "List Paragraph", "List Bullet", "List Number", "Quote"]
    target_styles += [f"Heading {i}" for i in range(1, 10)]
    for name in target_styles:
        try:
            style = doc.styles[name]
            if style and getattr(style, "font", None):
                style.font.color.rgb = RGBColor(0, 0, 0)
        except KeyError:
            pass


# ------------------ Helpers DOCX ------------------

def add_math(paragraph, latex: str):
    """
    Inserta una ecuación LaTeX como OMML (editable en Word).
    Si algo falla, hace fallback a texto con fuente Cambria Math.
    """
    try:
        mathml = latex2mathml(latex)
        omml = mathml2omml.convert(mathml)
        omml_element = ET.fromstring(omml)
        # Adjuntamos el nodo OMML directamente al párrafo
        paragraph._element.append(omml_element)
    except Exception:
        run = paragraph.add_run(latex)
        run.font.name = "Cambria Math"


def add_hrule(doc: Document):
    """Añade una regla horizontal simple (borde inferior del párrafo)."""
    p = doc.add_paragraph()
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set("w:val", "single")
    bottom.set("w:sz", "6")
    bottom.set("w:space", "1")
    bottom.set("w:color", "auto")
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_image_paragraph(doc: Document, img_bytes: bytes):
    """Inserta una imagen ajustada al ancho útil de la página."""
    section = doc.sections[-1]
    usable_width_emu = section.page_width - section.left_margin - section.right_margin
    EMUS_PER_INCH = 914400
    usable_width_in = float(usable_width_emu) / EMUS_PER_INCH

    stream = io.BytesIO(img_bytes)
    try:
        with PILImage.open(io.BytesIO(img_bytes)) as im:
            width_px, _ = im.size
            dpi_x = im.info.get("dpi", (96, 96))[0] or 96
            width_in = width_px / float(dpi_x)
            scale = min(1.0, usable_width_in / width_in) if width_in else 1.0
            new_width_in = max(0.1, width_in * scale)
            p = doc.add_paragraph()
            run = p.add_run()
            stream.seek(0)
            run.add_picture(stream, width=Inches(new_width_in))
            return
    except Exception:
        pass

    p = doc.add_paragraph()
    run = p.add_run()
    stream.seek(0)
    run.add_picture(stream, width=Inches(usable_width_in))


def add_hyperlink(paragraph, text: str, url: str):
    """Inserta un enlace clicable sencillo (subrayado, azul)."""
    try:
        part = paragraph.part
        r_id = part.relate_to(
            url,
            reltype="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", r_id)

        new_run = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")

        u = OxmlElement("w:u")
        u.set("w:val", "single")
        rPr.append(u)

        color = OxmlElement("w:color")
        color.set("w:val", "0000FF")
        rPr.append(color)

        new_run.append(rPr)
        t = OxmlElement("w:t")
        t.text = text
        new_run.append(t)
        hyperlink.append(new_run)
        paragraph._element.append(hyperlink)
    except Exception:
        paragraph.add_run(text)


# ------------------ Markdown → DOCX ------------------

IMG_INLINE_RE = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')

def split_text_math(text: str) -> List[Dict[str, str]]:
    """
    Divide un string en segmentos de texto y math.
    Detecta $$...$$ (bloque) y $...$ (inline) de forma segura/simple.
    Devuelve lista de dicts: {"type": "text"|"math_inline"|"math_block", "value": "..."}.
    """
    out: List[Dict[str, str]] = []
    i = 0
    n = len(text)
    while i < n:
        # bloque $$...$$
        if text.startswith("$$", i):
            j = text.find("$$", i + 2)
            if j != -1:
                formula = text[i + 2:j]
                out.append({"type": "math_block", "value": formula})
                i = j + 2
                continue
        # inline $...$
        if text[i] == "$":
            j = i + 1
            while j < n:
                if text[j] == "\\":
                    j += 2
                    continue
                if text[j] == "$":
                    formula = text[i + 1:j]
                    out.append({"type": "math_inline", "value": formula})
                    i = j + 1
                    break
                j += 1
            else:
                # no se cerró: lo tratamos como texto normal
                out.append({"type": "text", "value": text[i]})
                i += 1
            continue
        # texto normal
        out.append({"type": "text", "value": text[i]})
        i += 1

    # compactar textos contiguos
    compact: List[Dict[str, str]] = []
    for seg in out:
        if seg["type"] == "text":
            if compact and compact[-1]["type"] == "text":
                compact[-1]["value"] += seg["value"]
            else:
                compact.append(seg)
        else:
            compact.append(seg)
    return compact


def render_inline_text(paragraph, content: str):
    """
    Render básico de énfasis (***, **, *, ~~) y código inline dentro de un párrafo,
    además de insertar ecuaciones inline cuando haya $...$.
    """
    # Partimos primero por matemáticas
    segments = split_text_math(content)

    # Pequeño parser de énfasis para segmentos de texto
    emphasis_re = re.compile(r'(\*\*\*.+?\*\*\*|\*\*.+?\*\*|~~.+?~~|\*(?!\*)(.+?)(?<!\*)\*|`[^`]+`)')

    for seg in segments:
        if seg["type"] == "math_inline":
            add_math(paragraph, seg["value"])
            continue
        if seg["type"] == "math_block":
            # en inline, mostramos como ecuación centrada separada
            p = paragraph._p.getparent().add_p()
            para = paragraph._parent.add_paragraph()  # mantener API estable
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_math(para, seg["value"])
            continue

        # texto normal con énfasis
        text = seg["value"]
        pos = 0
        for m in emphasis_re.finditer(text):
            if m.start() > pos:
                paragraph.add_run(text[pos:m.start()])

            token = m.group(0)
            if token.startswith("***") and token.endswith("***"):
                r = paragraph.add_run(token[3:-3])
                r.bold = True
                r.italic = True
            elif token.startswith("**") and token.endswith("**"):
                r = paragraph.add_run(token[2:-2])
                r.bold = True
            elif token.startswith("~~") and token.endswith("~~"):
                r = paragraph.add_run(token[2:-2])
                r.strike = True
            elif token.startswith("`") and token.endswith("`"):
                r = paragraph.add_run(token[1:-1])
                r.font.name = "Consolas"
                r.font.size = Pt(10)
            else:
                # *cursiva*
                core = token[1:-1]
                r = paragraph.add_run(core)
                r.italic = True
            pos = m.end()

        if pos < len(text):
            paragraph.add_run(text[pos:])


def images_from_payload(form_images_field: Optional[str]) -> Dict[str, bytes]:
    """
    Convierte el campo form 'images' (JSON con [{id, image_base64}]) en dict id->bytes.
    """
    result: Dict[str, bytes] = {}
    if not form_images_field:
        return result
    try:
        arr = json.loads(form_images_field)
        if isinstance(arr, list):
            for obj in arr:
                img_id = obj.get("id")
                img_b64 = obj.get("image_base64")
                if img_id and img_b64:
                    try:
                        result[img_id] = base64.b64decode(img_b64)
                    except Exception:
                        pass
    except Exception:
        pass
    return result


def build_doc_from_markdown(md_text: str, images: Dict[str, bytes]) -> Document:
    """
    Renderiza Markdown a DOCX con:
    - headings, párrafos, blockquotes, listas, HR, tablas
    - imágenes inline ![alt](id) -> usa images[id]
    - LaTeX $...$ y $$...$$ como ecuaciones editables
    """
    doc = Document()
    force_styles_black(doc)

    # Marcadores especiales de página/separador en el texto
    md_text = md_text.replace("[NUEVA PÁGINA]", "\n\n<NUEVA_PAGINA/>\n\n")
    # "— separadores —" (según ejemplo del usuario) -> HR
    md_text = re.sub(r"(?im)^---\s*separadores\s*---\s*$", "\n\n<HR/>\n\n", md_text)

    md = MarkdownIt("commonmark").enable("table").enable("strikethrough")
    tokens = md.parse(md_text)

    list_style_stack: List[str] = []  # "bullet" | "ordered"

    # Para tablas
    current_table: Optional[List[List[str]]] = None
    current_row: Optional[List[str]] = None
    collecting_cell_text: Optional[str] = None

    for tok in tokens:
        ttype = tok.type

        # ----- Page break / HR marcados previamente -----
        if ttype == "inline" and tok.content:
            if "<NUEVA_PAGINA/>" in tok.content.strip():
                doc.add_page_break()
                continue
            if "<HR/>" in tok.content.strip():
                add_hrule(doc)
                continue

        # ----- Heading -----
        if ttype == "heading_open":
            level = int(tok.tag[1])
            p = doc.add_paragraph()
            p.style = f"Heading {min(max(level,1),9)}"
            continue
        if ttype == "heading_close":
            continue
        if ttype == "inline" and tok.level > 0 and tok.map and tokens[tokens.index(tok)-1].type == "heading_open":
            # contenido del heading
            p = doc.paragraphs[-1]
            render_inline_text(p, tok.content)
            continue

        # ----- Horizontal rule -----
        if ttype == "hr":
            add_hrule(doc)
            continue

        # ----- Blockquote -----
        if ttype == "blockquote_open":
            p = doc.add_paragraph()
            p.style = "Quote"
            continue
        if ttype == "blockquote_close":
            continue
        if ttype == "inline" and tokens[tokens.index(tok)-1].type == "blockquote_open":
            p = doc.paragraphs[-1]
            render_inline_text(p, tok.content)
            continue

        # ----- Lists -----
        if ttype == "bullet_list_open":
            list_style_stack.append("bullet")
            continue
        if ttype == "ordered_list_open":
            list_style_stack.append("ordered")
            continue
        if ttype == "bullet_list_close" or ttype == "ordered_list_close":
            if list_style_stack:
                list_style_stack.pop()
            continue
        if ttype == "list_item_open":
            style = "List Number" if (list_style_stack and list_style_stack[-1] == "ordered") else "List Bullet"
            p = doc.add_paragraph(style=style)
            continue
        if ttype == "list_item_close":
            continue
        if ttype == "inline" and tokens[tokens.index(tok)-1].type == "list_item_open":
            p = doc.paragraphs[-1]
            text = tok.content

            # checklists simples: [-] [x] al inicio
            m = re.match(r'^\s*\[([ xX])\]\s+(.*)$', text)
            if m:
                chk = "☑ " if m.group(1).lower() == "x" else "☐ "
                render_inline_text(p, chk + m.group(2))
            else:
                render_inline_text(p, text)
            continue

        # ----- Tables -----
        if ttype == "table_open":
            current_table = []
            continue
        if ttype == "tr_open":
            current_row = []
            continue
        if ttype in ("td_open", "th_open"):
            collecting_cell_text = ""
            continue
        if ttype == "inline" and collecting_cell_text is not None:
            collecting_cell_text += tok.content
            continue
        if ttype in ("td_close", "th_close"):
            current_row.append(collecting_cell_text or "")
            collecting_cell_text = None
            continue
        if ttype == "tr_close":
            if current_table is not None:
                current_table.append(current_row or [])
            current_row = None
            continue
        if ttype == "table_close":
            if current_table:
                cols = max((len(r) for r in current_table), default=0)
                table = doc.add_table(rows=len(current_table), cols=max(cols, 1))
                table.style = "Table Grid"
                for i, row in enumerate(current_table):
                    for j in range(cols):
                        table.cell(i, j).text = (row[j] if j < len(row) else "").strip()
            current_table = None
            continue

        # ----- Paragraphs -----
        if ttype == "paragraph_open":
            p = doc.add_paragraph()
            continue
        if ttype == "inline" and tokens[tokens.index(tok)-1].type == "paragraph_open":
            p = doc.paragraphs[-1]

            # Procesar imágenes inline ![](id) contra el mapa 'images'
            text = tok.content
            pos = 0
            for m in IMG_INLINE_RE.finditer(text):
                if m.start() > pos:
                    render_inline_text(p, text[pos:m.start()])
                img_id = (m.group(2) or "").strip()
                blob = images.get(img_id)
                if blob:
                    add_image_paragraph(doc, blob)
                    p = doc.add_paragraph()  # nueva línea tras imagen
                pos = m.end()
            if pos < len(text):
                render_inline_text(p, text[pos:])
            continue
        if ttype == "paragraph_close":
            continue

        # ----- Softbreak → nueva línea suave -----
        if ttype == "softbreak":
            doc.add_paragraph()
            continue

    force_styles_black(doc)
    return doc


# ------------------ Endpoint /docx (contrato intacto) ------------------

@app.post("/docx")
def make_docx():
    """
    Acepta:
    - application/json { markdown?: str, text?: str, images?: [{id, image_base64}], output_name?: str }
    - multipart/form-data con los mismos campos (como lo mandas desde n8n):
        - markdown (string)
        - images (string JSON de array)
        - output_name (string)
        - files opcionales (request.files) para imágenes adicionales
    Devuelve: .docx (attachment)
    """
    # 1) JSON
    data = request.get_json(silent=True)
    if data and isinstance(data, dict) and ("markdown" in data or "text" in data):
        base_name = data.get("output_name") or data.get("filename") or "output"
        if not base_name.lower().endswith(".docx"):
            base_name += ".docx"

        images_map: Dict[str, bytes] = {}
        if "images" in data and isinstance(data["images"], list):
            for img_obj in data["images"]:
                img_id = img_obj.get("id")
                img_b64 = img_obj.get("image_base64")
                if img_id and img_b64:
                    try:
                        images_map[img_id] = base64.b64decode(img_b64)
                    except Exception:
                        pass

        md_text = data.get("markdown")
        plain_text = data.get("text")

        if md_text:
            doc = build_doc_from_markdown(md_text, images_map)
        else:
            doc = Document()
            force_styles_black(doc)
            p = doc.add_paragraph()
            render_inline_text(p, plain_text or "")

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=base_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    # 2) FORM-DATA (n8n)
    if request.form and ("markdown" in request.form or "text" in request.form):
        md_text = request.form.get("markdown")
        plain_text = request.form.get("text")
        base_name = request.form.get("output_name") or request.form.get("filename") or "output"
        if not base_name.lower().endswith(".docx"):
            base_name += ".docx"

        # images desde campo JSON
        images_map = images_from_payload(request.form.get("images"))

        # images subidas como ficheros
        for f in request.files.getlist("file"):
            if isinstance(f, FileStorage) and f.filename:
                images_map[f.filename] = f.read()

        if md_text:
            doc = build_doc_from_markdown(md_text, images_map)
        else:
            doc = Document()
            force_styles_black(doc)
            p = doc.add_paragraph()
            render_inline_text(p, plain_text or "")

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=base_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    return jsonify({"error": "Bad request"}), 400


# ------------------ Endpoint /merge (SIN TOCAR) ------------------

@app.post("/merge")
def merge_docx():
    data = request.get_json(silent=True)
    if not data:
        data = request.form.to_dict()

    if not data or "docs" not in data:
        return jsonify({"error": "Bad request: falta 'docs'"}), 400

    docs_field = data["docs"]

    if isinstance(docs_field, str):
        try:
            docs_dict = json.loads(docs_field)
        except Exception as e:
            return jsonify({"error": f"No se pudo parsear 'docs': {e}"}), 400
    elif isinstance(docs_field, dict):
        docs_dict = docs_field
    else:
        return jsonify({"error": "Formato de 'docs' no válido"}), 400

    if not docs_dict:
        return jsonify({"error": "No se recibió ningún documento para mergear"}), 400

    merged = None
    composer = None

    for key in sorted(docs_dict.keys(), key=lambda x: int(x)):
        b64 = docs_dict[key]
        try:
            content = base64.b64decode(b64)
            subdoc = Document(io.BytesIO(content))
        except Exception as e:
            return jsonify({"error": f"Error procesando doc {key}: {e}"}), 400

        if merged is None:
            merged = subdoc
            composer = Composer(merged)
        else:
            # Forzar nueva página entre documentos
            merged.add_section(WD_SECTION.NEW_PAGE)
            composer.append(subdoc)

    buf = io.BytesIO()
    merged.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name=data.get("output_name", "merged.docx"),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
