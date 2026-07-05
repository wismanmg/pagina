# Dockerfile para Render (servicio tipo Docker)
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render inyecta el puerto en la variable $PORT
CMD gunicorn --bind 0.0.0.0:$PORT app:app
