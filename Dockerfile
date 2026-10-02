FROM python:3.12-slim
WORKDIR /app
ENV PYTHONUNBUFFERED=1 NUMBA_CACHE_DIR=/tmp/numba DATA_DIR=/data
COPY backend/requirements.txt backend/requirements.txt
RUN pip install --no-cache-dir -r backend/requirements.txt
RUN mkdir -p /data
COPY backend backend
COPY frontend frontend
WORKDIR /app/backend
# $PORT diisi otomatis oleh Railway/Render; default 8000 untuk lokal
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"]
