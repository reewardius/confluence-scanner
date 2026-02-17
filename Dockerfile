FROM python:3.11-slim

# Install system dependencies: tesseract OCR + language data
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-eng \
    tesseract-ocr-rus \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Set tessdata path (already hardcoded in the script)
ENV TESSDATA_PREFIX=/usr/share/tesseract-ocr/5/tessdata/

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the script
COPY confluence.py .

# Output directory for results
RUN mkdir -p /output

ENTRYPOINT ["python", "confluence.py"]
