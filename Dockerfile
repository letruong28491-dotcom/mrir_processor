FROM python:3.11-slim

# Cài đặt các thư viện hệ thống cần thiết cho OpenCV, PyMuPDF và LibreOffice (để convert Excel sang PDF)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libgl1 \
    libglib2.0-0 \
    libreoffice \
    fonts-dejavu \
    fonts-liberation \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Sao chép và cài đặt các thư viện python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Sao chép toàn bộ mã nguồn vào container
COPY . .

# Tạo thư mục Input và Output
RUN mkdir -p Input Output

EXPOSE 8501

# Lệnh khởi chạy Streamlit
ENTRYPOINT ["streamlit", "run", "app_web.py", "--server.port=8501", "--server.address=0.0.0.0", "--server.enableCORS=false", "--server.enableXsrfProtection=false"]
