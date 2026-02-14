# Production-ready Dockerfile for DocMerge API
# Optimized for Railway deployment with LibreOffice support

FROM python:3.11-slim

# Set environment variables for Python optimization and LibreOffice
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    HOME=/home/appuser \
    XDG_RUNTIME_DIR=/tmp \
    SAL_USE_VCLPLUGIN=gen

# Install system dependencies
# LibreOffice for document processing
# poppler-utils for PDF to image conversion (required by pdf2image)
# Additional dependencies for image processing
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        poppler-utils \
        libpoppler-cpp-dev \
        libjpeg-dev \
        zlib1g-dev \
        libpng-dev \
        && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements first for better Docker layer caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create non-root user for security and ensure temp directories are writable
RUN useradd -m -u 1000 appuser && \
    chown -R appuser:appuser /app && \
    chmod 1777 /tmp && \
    mkdir -p /home/appuser/.config && \
    chown -R appuser:appuser /home/appuser

# Switch to non-root user
USER appuser

# Expose port 8080 (Railway default)
EXPOSE 8080

# Health check using built-in Python libraries
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8080/')" || exit 1

# Run the application
# Use 0.0.0.0 to bind to all interfaces (required for Railway)
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080", "--workers", "1"]

