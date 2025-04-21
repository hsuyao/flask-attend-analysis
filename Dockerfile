# Use Python 3.8 as the base image
FROM python:3.8-slim

# Install system dependencies, including LibreOffice and Redis
RUN apt-get update && \
    apt-get install -y libreoffice redis-server && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Generate version_info.txt during build
RUN if [ -d .git ]; then \
        echo "git-$(git rev-parse --short HEAD)-$(date -u +%Y%m%d%H%M%S)" > /app/version_info.txt; \
    else \
        echo "custom-$(date -u +%Y%m%d%H%M%S)" > /app/version_info.txt; \
    fi

# Expose Flask application port
EXPOSE 5000

# Start Gunicorn and Redis
CMD redis-server --port 6379 & gunicorn --bind 0.0.0.0:5000 --workers=1 --threads=4 app:app
