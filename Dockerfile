# 使用 Python 3.8 作為基礎映像，與您的環境一致
FROM python:3.8-slim

# 安裝系統依賴，包括 LibreOffice（用於 Excel 處理）
RUN apt-get update && \
    apt-get install -y libreoffice redis-server && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# 設定工作目錄
WORKDIR /app

# 複製並安裝 Python 依賴
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 複製應用程式程式碼
COPY . .

# 暴露 Flask 應用端口
EXPOSE 5000

# 預設命令（將在 docker-compose.yml 中覆蓋）
CMD ["python", "app.py"]
