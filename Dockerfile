# Imagen base con Python
FROM python:3.11-slim

# Instalar pandoc y dependencias
RUN apt-get update && apt-get install -y pandoc && rm -rf /var/lib/apt/lists/*

# Crear directorio de la app
WORKDIR /app

# Copiar requirements e instalar
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el código
COPY . .

# Puerto para Render
ENV PORT=5000

# Comando de arranque con Gunicorn
CMD ["gunicorn", "-b", "0.0.0.0:5000", "main:app"]
