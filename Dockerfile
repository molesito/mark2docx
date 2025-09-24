FROM python:3.11-slim

WORKDIR /app

# Dependencias del sistema para lxml y pillow
RUN apt-get update && apt-get install -y --no-install-recommends \
    libxml2-dev libxslt1-dev gcc build-essential \
    libjpeg62-turbo-dev zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Flask + gunicorn
ENV PYTHONUNBUFFERED=1
ENV PORT=8000

CMD ["gunicorn", "-w", "2", "-k", "gthread", "--threads", "4", "--timeout", "90", "main:app", "--bind", "0.0.0.0:8000"]
