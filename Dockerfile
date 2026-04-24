FROM python:3.12-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    FLASK_APP=app.py \
    FLASK_ENV=production \
    UPLOAD_FOLDER=/app/uploads \
    OUTPUT_FOLDER=/app/outputs

# Set work directory
WORKDIR /app

# Install system dependencies for pdf2docx compatibility (ARM64 compatible)
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    libffi-dev \
    libssl-dev \
    curl \
    libglib2.0-0 \
    libgomp1 \
    libsm6 \
    libxext6 \
    libxrender1 \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies (using domestic mirror for faster download)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# Uninstall and reinstall opencv-python for headless version
RUN pip uninstall -y opencv-python opencv-contrib-python && \
    pip install --no-cache-dir opencv-python-headless -i https://pypi.tuna.tsinghua.edu.cn/simple

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