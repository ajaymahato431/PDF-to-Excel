FROM python:3.11-slim

# Install system dependencies: Tesseract OCR, Nepali language pack, and Poppler utilities
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-nep \
    poppler-utils \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Set standard working directory
WORKDIR /app

# Copy requirements and install python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy source code
COPY pdf_to_excel.py .

# Environment defaults for container
ENV PYTHONUNBUFFERED=1 \
    TESSDATA_PREFIX=/usr/share/tesseract-ocr/5/tessdata

# Default entrypoint
ENTRYPOINT ["python", "pdf_to_excel.py"]
CMD ["--help"]
