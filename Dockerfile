FROM python:3.11-slim

# Evitar prompts de tzdata, etc.
ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Actualiza e instala LibreOffice + dependencias y algunas fuentes
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    fonts-dejavu \
    fonts-liberation \
    ca-certificates \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Crear directorio de app
WORKDIR /app

# Copiar requirements e instalar
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# Copiar el código
COPY main.py /app/main.py

# Puerto para Render
ENV PORT=8000
EXPOSE 8000

# Comando de producción con Gunicorn (Uvicorn workers)
# Puedes ajustar -w (workers) según plan de Render
CMD ["gunicorn", "-k", "uvicorn.workers.UvicornWorker", "-w", "2", "-b", "0.0.0.0:8000", "main:app"]
