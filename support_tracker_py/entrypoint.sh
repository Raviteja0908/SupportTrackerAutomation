#!/usr/bin/env bash
set -euo pipefail

if [[ "${EMIT_SCRIPT:-0}" == "1" ]]; then
  HOST_DIR="${HOST_SCRIPT_DIR:-/host}"
  if [[ ! -d "$HOST_DIR" ]]; then
    echo "[ERROR] Host script dir not found: $HOST_DIR"
    echo "Mount a host folder with: -v \"D:\\Support_Tracker\\Scripts:/host\""
    exit 1
  fi
  mkdir -p "$HOST_DIR"
  cp -f /app/tools/export_outlook.ps1 "$HOST_DIR/export_outlook.ps1"
  cp -f /app/tools/run_pipeline.ps1 "$HOST_DIR/run_pipeline.ps1"
  echo "[INFO] Scripts written to $HOST_DIR"
  echo "[INFO] Next run on host:"
  echo "powershell -ExecutionPolicy Bypass -File \"$HOST_DIR/run_pipeline.ps1\""
  exit 0
fi

exec python3 /app/main.py
