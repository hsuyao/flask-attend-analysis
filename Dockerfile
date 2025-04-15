# Use an official Python runtime as the base image
FROM python:3.8-slim

# Install LibreOffice for soffice
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set the working directory
WORKDIR /app

# Copy requirements.txt first to leverage Docker caching
COPY requirements.txt .

# Upgrade pip and install Python dependencies
RUN pip install --upgrade pip && \
    pip install -r requirements.txt && \
    pip cache purge

# Copy the rest of the application files
COPY . .

# 在构建时获取 commit ID 和日期，写入 version_info.txt
RUN echo "echo '$(date -u +%Y%m%d)" > /app/version_info.txt

# Create directories for session storage and database
RUN mkdir -p /app/sessions /app/db

# Command to run the application with database initialization
CMD ["sh", "-c", "python -c 'from database import init_database; init_database()' && python app.py"]
