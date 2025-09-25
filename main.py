import os
import subprocess
import uuid
from pathlib import Path
from flask import Flask, request, send_file, abort

app = Flask(__name__)
UPLOAD_DIR = Path("/tmp/uploads")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

@app.route("/")
def home():
    return "Servicio de conversión LaTeX → DOCX con ecuaciones Word. Sube tu .docx en /upload"

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file or not file.filename.endswith(".docx"):
        return abort(400, "Debes subir un archivo .docx")

    uid = uuid.uuid4().hex
    input_path = UPLOAD_DIR / f"{uid}.docx"
    md_path = UPLOAD_DIR / f"{uid}.md"
    output_path = UPLOAD_DIR / f"{uid}_out.docx"

    file.save(input_path)

    try:
        # Paso 1: DOCX -> Markdown (para que pandoc reconozca $...$)
        subprocess.run([
            "pandoc", str(input_path),
            "-f", "docx", "-t", "markdown",
            "-o", str(md_path)
        ], check=True)

        # Paso 2: Markdown -> DOCX (pandoc convierte a OMML)
        subprocess.run([
            "pandoc", str(md_path),
            "-f", "markdown", "-t", "docx",
            "-o", str(output_path)
        ], check=True)

    except subprocess.CalledProcessError:
        return abort(500, "Error al convertir con pandoc")

    return send_file(output_path, as_attachment=True, download_name="resultado.docx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
