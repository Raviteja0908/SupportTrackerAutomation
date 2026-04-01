#!/usr/bin/env bash
set -euo pipefail

if [[ "${EMIT_SCRIPT:-0}" == "1" ]]; then
  HOST_ROOT="${HOST_ROOT_DIR:-/hostroot}"
  HOST_DIR="${HOST_SCRIPT_DIR:-/host}"
  if [[ -d "$HOST_ROOT" ]]; then
    SCRIPTS_DIR="$HOST_ROOT/Scripts"
    PST_DIR="$HOST_ROOT/PstFiles"
    OUTPUT_DIR="$HOST_ROOT/DockerOutput"
    mkdir -p "$SCRIPTS_DIR" "$PST_DIR" "$OUTPUT_DIR"
    cp -f /app/tools/export_outlook.ps1 "$SCRIPTS_DIR/export_outlook.ps1"
    cp -f /app/tools/run_pipeline.ps1 "$SCRIPTS_DIR/run_pipeline.ps1"
    echo "[INFO] Scripts written to: $SCRIPTS_DIR"
    echo "[INFO] PST folder ready at: $PST_DIR"
    echo "[INFO] Output folder ready at: $OUTPUT_DIR"
    echo "[INFO] Next run on host:"
    echo "powershell -ExecutionPolicy Bypass -File \"<your-mounted-Support_Tracker>\\Scripts\\run_pipeline.ps1\""
    exit 0
  fi
  if [[ ! -d "$HOST_DIR" ]]; then
    echo "[ERROR] Host script dir not found: $HOST_DIR"
    echo "Mount a host root folder with: -v \"D:\\Support_Tracker:/hostroot\""
    echo "Or use the older scripts-only mount: -v \"D:\\Support_Tracker\\Scripts:/host\""
    exit 1
  fi
  mkdir -p "$HOST_DIR"
  cp -f /app/tools/support_tracker.ps1 "$HOST_DIR/support_tracker.ps1"
  echo "[INFO] Scripts written to: $HOST_DIR"
  echo "[INFO] Next run on host:"
  echo "powershell -ExecutionPolicy Bypass -File \".\\support_tracker.ps1\""
  exit 0
fi

exec python3 /app/main.py
