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
    libglapi-mesa \
    libgl1-mesa-dri \
    libsm6 \
    libxext6 \
    libxrender-dev \
    libgthread-2.0-0 \
    xvfb \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables for headless OpenCV
ENV DISPLAY=:99
ENV OPENCV_IO_ENABLE_OPENEXR=1

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Uninstall and reinstall opencv-python for headless version
RUN pip uninstall -y opencv-python opencv-contrib-python && \
    pip install --no-cache-dir opencv-python-headless

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

# Start virtual display and run the application
CMD ["/bin/bash", "-c", "Xvfb :99 -screen 0 1024x768x24 > /dev/null 2>&1 & python app.py"]