FROM python:3.11-slim

# نصب LibreOffice و فونت‌های مورد نیاز برای فارسی
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# کپی و نصب کتابخانه‌های پایتون
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# کپی تمام فایل‌های پروژه به ریشه /app
COPY . .

# ساخت پوشه فاکتورها
RUN mkdir -p /app/invoices

# اجرای فایل main.py که الان در ریشه قرار داره
CMD ["python", "main.py"]
