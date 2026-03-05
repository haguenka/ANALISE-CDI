FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive
WORKDIR /opt/app

RUN apt-get update && apt-get install -y --no-install-recommends \
    fluxbox \
    novnc \
    websockify \
    x11vnc \
    xvfb \
    fonts-dejavu-core \
    libasound2 \
    libdbus-1-3 \
    libegl1 \
    libfontconfig1 \
    libgl1 \
    libglib2.0-0 \
    libice6 \
    libnss3 \
    libopengl0 \
    libsm6 \
    libx11-6 \
    libx11-xcb1 \
    libxcb-cursor0 \
    libxcb-glx0 \
    libxcb-icccm4 \
    libxcb-image0 \
    libxcb-keysyms1 \
    libxcb-randr0 \
    libxcb-render-util0 \
    libxcb-render0 \
    libxcb-shape0 \
    libxcb-shm0 \
    libxcb-sync1 \
    libxcb-xfixes0 \
    libxcb-xinerama0 \
    libxcb-xkb1 \
    libxcomposite1 \
    libxcursor1 \
    libxdamage1 \
    libxext6 \
    libxfixes3 \
    libxi6 \
    libxkbcommon-x11-0 \
    libxkbcommon0 \
    libxrandr2 \
    libxrender1 \
    libxtst6 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt /opt/app/requirements.txt
RUN pip install --no-cache-dir -r /opt/app/requirements.txt

RUN mkdir -p /opt/app/novnc-web \
    && cp -r /usr/share/novnc/. /opt/app/novnc-web/

COPY app /opt/app/app
COPY web /opt/app/web
COPY start.sh /opt/app/start.sh

RUN cp /opt/app/web/index.html /opt/app/novnc-web/index.html \
    && chmod +x /opt/app/start.sh

EXPOSE 10000
CMD ["/opt/app/start.sh"]
