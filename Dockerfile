FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOST=0.0.0.0 \
    PORT=8765 \
    INVOICE_CLOUD=1 \
    INVOICE_DATA_DIR=/app/data \
    INVOICE_TEMPLATE=/app/发票模板.xlsx

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        poppler-utils \
        tesseract-ocr \
        tesseract-ocr-chi-sim \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY 发票模板.xlsx ./发票模板.xlsx
COPY invoice_system ./invoice_system

RUN mkdir -p /app/data/uploads /app/data/outputs
EXPOSE 8765

CMD ["python", "invoice_system/web_app.py"]
