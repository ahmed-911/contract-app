FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    fonts-dejavu \
    fonts-noto-core \
    fonts-noto-extra \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
