FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp

# Install LibreOffice Calc (headless Excel->PDF), CJK Traditional Chinese fonts, and OpenCV runtime libs
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    fonts-noto-cjk \
    fonts-arphic-ukai \
    libgl1 \
    libglib2.0-0 \
    && printf '<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "fonts.dtd">\n<fontconfig>\n  <alias><family>標楷體</family><prefer><family>AR PL UKai TW</family></prefer></alias>\n  <alias><family>DFKai-SB</family><prefer><family>AR PL UKai TW</family></prefer></alias>\n</fontconfig>\n' > /etc/fonts/local.conf \
    && fc-cache -fv \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5050

CMD ["sh", "-c", "waitress-serve --host=0.0.0.0 --port=${PORT:-5050} --threads=2 app:app"]

