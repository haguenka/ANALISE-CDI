#!/usr/bin/env bash
set -euo pipefail

export DISPLAY=:99
export CDI_DATA_DIR="${CDI_DATA_DIR:-/var/data}"
export XDG_RUNTIME_DIR="${XDG_RUNTIME_DIR:-/tmp/runtime-root}"
export QT_QPA_PLATFORM="${QT_QPA_PLATFORM:-xcb}"
export QT_X11_NO_MITSHM=1
export QT_OPENGL=software
export QT_QUICK_BACKEND=software
export LIBGL_ALWAYS_SOFTWARE=1
export MESA_LOADER_DRIVER_OVERRIDE=llvmpipe
export PYTHONUNBUFFERED=1

mkdir -p "${CDI_DATA_DIR}" "${XDG_RUNTIME_DIR}"
chmod 700 "${XDG_RUNTIME_DIR}"

Xvfb :99 -screen 0 1720x1080x24 -ac +extension GLX +render -noreset >/tmp/xvfb.log 2>&1 &
fluxbox >/tmp/fluxbox.log 2>&1 &
x11vnc -display :99 -forever -shared -rfbport 5900 -nopw -localhost >/tmp/x11vnc.log 2>&1 &

for _ in $(seq 1 30); do
  if xdpyinfo -display "${DISPLAY}" >/dev/null 2>&1; then
    break
  fi
  sleep 1
done

if ! xdpyinfo -display "${DISPLAY}" >/dev/null 2>&1; then
  echo "X display ${DISPLAY} did not become ready"
  exit 1
fi

python -X faulthandler -u /opt/app/app/analise_tempo_atendimento_cdi.py 2>&1 | tee /tmp/cdi_app.log &
APP_PID=$!

websockify --web=/opt/app/novnc-web/ 6080 localhost:5900 >/tmp/novnc.log 2>&1 &
NOVNC_PID=$!

python /opt/app/upload_server.py >/tmp/upload_server.log 2>&1 &
UPLOAD_PID=$!

sed "s/__PORT__/${PORT:-10000}/g" /opt/app/nginx.conf.template >/tmp/nginx.conf
nginx -c /tmp/nginx.conf -g 'daemon off;' >/tmp/nginx.log 2>&1 &
NGINX_PID=$!

cleanup() {
  kill "${APP_PID}" "${NOVNC_PID}" "${UPLOAD_PID}" "${NGINX_PID}" 2>/dev/null || true
}

trap cleanup EXIT
wait -n "${APP_PID}" "${NOVNC_PID}" "${UPLOAD_PID}" "${NGINX_PID}"
