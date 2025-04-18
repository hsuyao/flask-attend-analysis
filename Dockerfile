# 使用 Python 3.8 作為基礎映像
FROM python:3.8-slim

# 安裝系統依賴，包括 LibreOffice、Redis 和 supervisord
RUN apt-get update && \
    apt-get install -y libreoffice redis-server supervisor && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# 設定工作目錄
WORKDIR /app

# 複製並安裝 Python 依賴
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 複製應用程式程式碼
COPY . .

# 複製 supervisord 配置文件
COPY supervisord.conf /etc/supervisor/conf.d/supervisord.conf

# 暴露 Flask 應用端口
EXPOSE 5000

# 使用 supervisord 作為容器啟動命令
CMD ["/usr/bin/supervisord", "-c", "/etc/supervisor/conf.d/supervisord.conf"]
