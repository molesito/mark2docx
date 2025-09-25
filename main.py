import os
import shutil
import subprocess
import tempfile
from datetime import datetime
from typing import Optional

from fastapi import FastAPI, HTTPException, Response
from pydantic import BaseModel

app = FastAPI(title="HTML to DOCX", version="1.0.0")


class HtmlPayload(BaseModel):
    html: str
    filename: Optional[str] = None  # opcional: nombre del archivo de salida


def _convert_html_to_docx(html_str: str, desired_name: Optional[str]) -> bytes:
    if not html_str or not html_str.strip():
        raise ValueError("El campo 'html' está vacío.")

    # Carpeta temporal aislada por petición
    with tempfile.TemporaryDirectory() as tmpdir:
        html_path = os.path.join(tmpdir, "input.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_str)

        # LibreOffice generará un .docx con el mismo basename
        # Ejemplo de comando:
        # soffice --headless --convert-to docx --outdir /tmp/tmpabcd input.html
        cmd = [
            "soffice",
            "--headless",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "--norestore",
            "--invisible",
            "--convert-to",
            "docx:MS Word 2007 XML",
            "--outdir",
            tmpdir,
            html_path,
        ]
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=tmpdir,
        )

        if proc.returncode != 0:
            # Incluir algo de contexto de error
            raise RuntimeError(
                f"Error en LibreOffice (code {proc.returncode}). "
                f"STDOUT: {proc.stdout}\nSTDERR: {proc.stderr}"
            )

        # Buscar el .docx generado
        # Por defecto, será input.docx
        out_path = os.path.join(tmpdir, "input.docx")
        if not os.path.exists(out_path):
            # En caso de cambios de nombre raros (no debería), buscar el primer .docx
            candidates = [p for p in os.listdir(tmpdir) if p.lower().endswith(".docx")]
            if not candidates:
                raise RuntimeError("No se encontró el archivo DOCX de salida.")
            out_path = os.path.join(tmpdir, candidates[0])

        with open(out_path, "rb") as f:
            data = f.read()

        # Si el usuario quiere un nombre concreto, lo gestionamos en la cabecera, no aquí
        return data


@app.get("/health")
def health():
    return {"status": "ok", "timestamp": datetime.utcnow().isoformat() + "Z"}


@app.post("/to-docx")
def to_docx(payload: HtmlPayload):
    try:
        data = _convert_html_to_docx(payload.html, payload.filename)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve)) from ve
    except RuntimeError as re:
        raise HTTPException(status_code(500), detail=str(re)) from re
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error inesperado: {e}") from e

    # Nombre final sugerido
    fname = payload.filename.strip() if payload.filename else "document.docx"
    if not fname.lower().endswith(".docx"):
        fname += ".docx"

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )
