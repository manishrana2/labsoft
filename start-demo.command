#!/bin/bash
set -e

cd "$(dirname "$0")"

echo "=========================================="
echo "Labsoft Demo Launcher (macOS)"
echo "=========================================="

echo
if ! command -v node >/dev/null 2>&1; then
  echo "[ERROR] Node.js not found."
  echo "Install Node.js LTS from https://nodejs.org/"
  exit 1
fi

echo "Installing dependencies (first run may take a while)..."
npm install

echo
echo "Starting Labsoft demo on http://localhost:5173 ..."
echo "Keep this Terminal window open while demo is running."
echo "Press Ctrl+C to stop."
echo
open "http://localhost:5173" >/dev/null 2>&1 || true
npm run dev:full
