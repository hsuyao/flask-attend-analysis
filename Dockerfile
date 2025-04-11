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
# 如果没有 git 环境，则使用默认值
RUN echo "$(git rev-parse HEAD 2>/dev/null || echo 'Unknown')-$(date -u +%Y%m%d)" > /app/version_info.txt || echo "Unknown-Unknown" > /app/version_info.txt

# Create a directory for session storage
RUN mkdir -p /app/sessions

# Expose the port
EXPOSE 5000

# Command to run the application
CMD ["python", "app.py"]
