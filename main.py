import os
import tempfile
import subprocess
from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pydantic import BaseModel

app = FastAPI(title="MD→DOCX Converter")


class Payload(BaseModel):
    markdown: str
    filename: str = "document.docx"


def md_to_docx(md_text: str, out_path: str):
    if not md_text.strip():
        raise ValueError("El contenido Markdown está vacío.")

    cmd = [
        "pandoc",
        "-f", "markdown+tex_math_dollars",
        "-t", "docx",
        "--standalone",
        "--wrap=preserve",
        "-o", out_path,
        "-"
    ]
    subprocess.run(cmd, input=md_text.encode("utf-8"), check=True)


@app.post("/convert")
def convert(payload: Payload, background_tasks: BackgroundTasks):
    # Archivo temporal de salida
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        out_path = tmp.name

    try:
        md_to_docx(payload.markdown, out_path)
    except Exception as e:
        if os.path.exists(out_path):
            os.remove(out_path)
        raise HTTPException(status_code=400, detail=str(e))

    # Limpiar después de enviar
    background_tasks.add_task(lambda p: os.path.exists(p) and os.remove(p), out_path)

    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=payload.filename,
    )
