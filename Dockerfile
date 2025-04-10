# Use an official Python runtime as the base image
FROM python:3.8-slim

# Install LibreOffice
RUN apt-get update && apt-get install -y libreoffice && apt-get clean

# Set the working directory
WORKDIR /app

# Copy the application files
COPY . .

# Install Python dependencies
RUN pip install -r requirements.txt

# Expose the port
# EXPOSE 5000

# Command to run the application
CMD ["python", "app.py"]
