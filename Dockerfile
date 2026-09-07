FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=5050

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source code
COPY . .

EXPOSE 5050

# Run with Gunicorn using dynamic PORT
CMD exec gunicorn app:app --bind 0.0.0.0:${PORT:-5050} --workers 2 --threads 4 --timeout 120
