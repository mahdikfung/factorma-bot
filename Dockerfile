FROM python:3.11-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Create directory for invoices if not exists
RUN mkdir -p /app/invoices

# Expose port (not needed for bot, but good practice)
# EXPOSE 8080

# Run the bot
CMD ["python", "bot/main.py"]