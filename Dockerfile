FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=5050 \
    DEBIAN_FRONTEND=noninteractive

WORKDIR /app

# Install LibreOffice and fonts for high-fidelity DOCX to PDF conversion on Linux
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-calc \
    fonts-dejavu-core \
    fonts-liberation \
    fonts-noto-core \
    chromium \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source code
COPY . .

EXPOSE 5050

# Run with Gunicorn using dynamic PORT (single worker with threads to maximize RAM on 512MB cloud instances)
CMD exec gunicorn app:app --bind 0.0.0.0:${PORT:-5050} --workers 1 --threads 8 --timeout 300
