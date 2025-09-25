FROM python:3.11-slim

# Instalar pandoc
RUN apt-get update && apt-get install -y pandoc && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=5000

CMD ["gunicorn", "-b", "0.0.0.0:5000", "main:app"]
