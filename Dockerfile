FROM python:3.12-slim
# libcairo2: vẽ lại hình (SVG -> PNG) · fonts-liberation: phông tương thích Times New Roman
# libreoffice-writer-nogui: xem trước trang A4 của đề (tuỳ chọn, ~350 MB). Bỏ qua: docker build --build-arg WITH_LIBREOFFICE=0 .
ARG WITH_LIBREOFFICE=1
RUN apt-get update \
 && apt-get install -y --no-install-recommends libcairo2 fonts-liberation fonts-dejavu-core \
 && if [ "$WITH_LIBREOFFICE" = "1" ]; then apt-get install -y --no-install-recommends libreoffice-writer-nogui; fi \
 && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
# Chạy dưới user không phải root (giảm rủi ro nếu có lỗ hổng)
RUN useradd -m -u 10001 appuser && chown -R appuser:appuser /app
ENV DATA_DIR=/app/data PORT=8000
EXPOSE 8000
USER appuser
# Kiểm tra sức khỏe qua endpoint /api/health (không cần đăng nhập)
HEALTHCHECK --interval=60s --timeout=10s --start-period=90s --retries=3 \
  CMD python -c "import os,urllib.request;urllib.request.urlopen('http://127.0.0.1:'+os.environ.get('PORT','8000')+'/api/health',timeout=8)" || exit 1
CMD ["sh","-c","uvicorn main:app --host 0.0.0.0 --port ${PORT}"]
