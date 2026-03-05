#!/usr/bin/env bash
set -euo pipefail

export DISPLAY=:99
export CDI_DATA_DIR="${CDI_DATA_DIR:-/var/data}"
export XDG_RUNTIME_DIR="${XDG_RUNTIME_DIR:-/tmp/runtime-root}"
export QT_QPA_PLATFORM="${QT_QPA_PLATFORM:-xcb}"
export QT_X11_NO_MITSHM=1
export PYTHONUNBUFFERED=1

mkdir -p "${CDI_DATA_DIR}" "${XDG_RUNTIME_DIR}"
chmod 700 "${XDG_RUNTIME_DIR}"

Xvfb :99 -screen 0 1720x1080x24 -ac +extension GLX +render -noreset >/tmp/xvfb.log 2>&1 &
fluxbox >/tmp/fluxbox.log 2>&1 &
x11vnc -display :99 -forever -shared -rfbport 5900 -nopw -localhost >/tmp/x11vnc.log 2>&1 &

python /opt/app/app/analise_tempo_atendimento_cdi.py >/tmp/cdi_app.log 2>&1 &
APP_PID=$!

websockify --web=/opt/app/novnc-web/ "${PORT:-10000}" localhost:5900 >/tmp/novnc.log 2>&1 &
NOVNC_PID=$!

cleanup() {
  kill "${APP_PID}" "${NOVNC_PID}" 2>/dev/null || true
}

trap cleanup EXIT
wait -n "${APP_PID}" "${NOVNC_PID}"
