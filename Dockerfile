FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp

# Install LibreOffice Calc (headless Excel->PDF), CJK Traditional Chinese fonts, and OpenCV runtime libs
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    fonts-noto-cjk \
    fonts-arphic-ukai \
    fonts-arphic-uming \
    libgl1 \
    libglib2.0-0 \
    && printf '<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "fonts.dtd">\n<fontconfig>\n  <match target="pattern">\n    <test qual="any" name="family"><string>標楷體</string></test>\n    <edit name="family" mode="assign" binding="same"><string>AR PL UKai TW</string></edit>\n  </match>\n  <match target="pattern">\n    <test qual="any" name="family"><string>DFKai-SB</string></test>\n    <edit name="family" mode="assign" binding="same"><string>AR PL UKai TW</string></edit>\n  </match>\n  <match target="pattern">\n    <test qual="any" name="family"><string>新細明體</string></test>\n    <edit name="family" mode="assign" binding="same"><string>AR PL UMing TW</string></edit>\n  </match>\n  <match target="pattern">\n    <test qual="any" name="family"><string>PMingLiU</string></test>\n    <edit name="family" mode="assign" binding="same"><string>AR PL UMing TW</string></edit>\n  </match>\n</fontconfig>\n' > /etc/fonts/local.conf \
    && fc-cache -fv \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5050

CMD ["sh", "-c", "waitress-serve --host=0.0.0.0 --port=${PORT:-5050} --threads=2 app:app"]

