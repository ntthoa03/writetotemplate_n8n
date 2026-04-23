# Sử dụng Python bản nhẹ (slim) để giảm dung lượng
FROM python:3.11-slim

# Cài đặt các thư viện hệ thống cần thiết cho việc xử lý file Excel và mạng
RUN apt-get update && apt-get install -y \
    gcc \
    python3-dev \
    && rm -rf /var/lib/apt/lists/*

# Thiết lập thư mục làm việc trong container
WORKDIR /app

# Copy file danh sách thư viện vào trước để tận dụng cache của Docker
COPY requirements.txt .

# Cài đặt các thư viện Python
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ code vào container
COPY . .

# Mở port 8000
EXPOSE 8000

# Lệnh chạy server (Sử dụng 0.0.0.0 để n8n bên ngoài có thể gọi vào)
CMD ["python", "main.py"]