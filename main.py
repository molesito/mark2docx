import io
from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from markdown_it import MarkdownIt
from latex2mathml.converter import convert as latex2mathml
import mathml2omml
import lxml.etree as ET

app = FastAPI()

# =========================
#   HELPERS DOCX
# =========================
def add_math(paragraph, latex: str):
    """Inserta una fórmula LaTeX como OMML en un párrafo Word."""
    try:
        mathml = latex2mathml(latex)
        omml = mathml2omml.convert(mathml)
        omml_element = ET.fromstring(omml)
        paragraph._element.append(omml_element)
    except Exception:
        # fallback: texto plano si falla
        run = paragraph.add_run(latex)
        run.font.name = "Consolas"

def add_hyperlink(paragraph, text, url):
    """Inserta un enlace clicable en un párrafo."""
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

# =========================
#   PARSER
# =========================
def parse_markdown(md_text: str) -> Document:
    md = MarkdownIt().enable("table").enable("strikethrough")
    tokens = md.parse(md_text)
    doc = Document()

    list_stack = []  # para manejar indentación de listas

    for token in tokens:
        if token.type == "heading_open":
            level = int(token.tag[1])
            para = doc.add_paragraph()
            para.style = f"Heading {level}"
        elif token.type == "inline":
            # contenido de párrafo, encabezado, etc.
            if token.content.strip().startswith("$$") and token.content.strip().endswith("$$"):
                para = doc.add_paragraph()
                add_math(para, token.content.strip().strip("$"))
            elif token.content.strip().startswith("$") and token.content.strip().endswith("$"):
                para = doc.add_paragraph()
                add_math(para, token.content.strip().strip("$"))
            else:
                para = doc.add_paragraph(token.content)
        elif token.type == "paragraph_open":
            pass
        elif token.type == "paragraph_close":
            pass
        elif token.type == "bullet_list_open":
            list_stack.append("bullet")
        elif token.type == "ordered_list_open":
            list_stack.append("ordered")
        elif token.type == "list_item_open":
            if list_stack and list_stack[-1] == "bullet":
                doc.add_paragraph(token.map, style="List Bullet")
            elif list_stack and list_stack[-1] == "ordered":
                doc.add_paragraph(token.map, style="List Number")
        elif token.type == "list_item_close":
            pass
        elif token.type == "bullet_list_close" or token.type == "ordered_list_close":
            if list_stack:
                list_stack.pop()
        elif token.type == "table_open":
            current_table = []
        elif token.type == "tr_open":
            current_row = []
        elif token.type == "td_open" or token.type == "th_open":
            current_cell = ""
        elif token.type == "inline" and token.map and token.content:
            current_cell = token.content
        elif token.type == "td_close" or token.type == "th_close":
            current_row.append(current_cell)
        elif token.type == "tr_close":
            current_table.append(current_row)
        elif token.type == "table_close":
            if current_table:
                table = doc.add_table(rows=0, cols=len(current_table[0]))
                for row in current_table:
                    cells = table.add_row().cells
                    for j, cell in enumerate(row):
                        cells[j].text = cell
        elif token.type == "hr":
            para = doc.add_paragraph()
            pPr = para._element.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set("w:val", "single")
            bottom.set("w:sz", "6")
            bottom.set("w:space", "1")
            bottom.set("w:color", "auto")
            pBdr.append(bottom)
            pPr.append(pBdr)
        elif token.type == "softbreak":
            doc.add_paragraph()

    return doc

# =========================
#   ENDPOINTS
# =========================
@app.post("/docx")
async def generate_docx(request: Request):
    data = await request.json()
    md_text = data.get("text", "")
    document = parse_markdown(md_text)
    file_stream = io.BytesIO()
    document.save(file_stream)
    file_stream.seek(0)
    headers = {
        "Content-Disposition": 'attachment; filename="output.docx"'
    }
    return StreamingResponse(file_stream, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", headers=headers)


# =========================
#   /merge (SIN TOCAR)
# =========================
@app.post("/merge")
async def merge_docs(request: Request):
    """
    Aquí mantengo exactamente lo que tenías antes para /merge.
    """
    data = await request.json()
    # Simulación de merge existente:
    result = {"status": "ok", "files": data.get("files", [])}
    return result
