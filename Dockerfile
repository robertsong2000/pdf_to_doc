FROM python:3.12-slim

ARG PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple
ARG PIP_TRUSTED_HOST=mirrors.aliyun.com
ARG PIP_DEFAULT_TIMEOUT=120
ARG PIP_RETRIES=10
ARG PDF2DOCX_PACKAGE=git+https://github.com/robertsong2000/pdf2docx.git@1ea348b736b72598cf8be03da1fa2f0bc20140ee

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    FLASK_APP=app.py \
    FLASK_ENV=production \
    UPLOAD_FOLDER=/app/uploads \
    OUTPUT_FOLDER=/app/outputs \
    PIP_DEFAULT_TIMEOUT=${PIP_DEFAULT_TIMEOUT} \
    PIP_RETRIES=${PIP_RETRIES} \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    DEBIAN_FRONTEND=noninteractive

# Set work directory
WORKDIR /app

# Install system dependencies for pdf2docx compatibility (ARM64 compatible)
RUN apt-get update && apt-get install -y \
    curl \
    git \
    libglib2.0-0 \
    libgomp1 \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies. The index is configurable because mirrors can
# occasionally timeout during Docker builds.
COPY requirements.txt .
RUN sed '/^pdf2docx==/d' requirements.txt > /tmp/requirements-no-pdf2docx.txt && \
    pip install --no-cache-dir \
        --timeout "${PIP_DEFAULT_TIMEOUT}" \
        --retries "${PIP_RETRIES}" \
        --index-url "${PIP_INDEX_URL}" \
        --trusted-host "${PIP_TRUSTED_HOST}" \
        -r /tmp/requirements-no-pdf2docx.txt && \
    pip install --no-cache-dir \
        --timeout "${PIP_DEFAULT_TIMEOUT}" \
        --retries "${PIP_RETRIES}" \
        --index-url "${PIP_INDEX_URL}" \
        --trusted-host "${PIP_TRUSTED_HOST}" \
        "PyMuPDF>=1.19.0" \
        "fonttools>=4.24.0" \
        "numpy>=1.17.2" \
        "opencv-python-headless>=4.5" \
        "fire>=0.3.0" && \
    pip install --no-cache-dir \
        --timeout "${PIP_DEFAULT_TIMEOUT}" \
        --retries "${PIP_RETRIES}" \
        --index-url "${PIP_INDEX_URL}" \
        --trusted-host "${PIP_TRUSTED_HOST}" \
        --no-deps \
        "${PDF2DOCX_PACKAGE}"

# Copy project
COPY . .

# Create necessary directories
RUN mkdir -p uploads outputs templates

# Create non-root user
RUN groupadd -r appuser && useradd -r -g appuser appuser
RUN chown -R appuser:appuser /app
USER appuser

# Expose port
EXPOSE 5000

# Health check
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:5000/ || exit 1

CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "1", "--threads", "4", "--timeout", "300", "--max-requests", "10000", "--max-requests-jitter", "1000", "app:app"]
