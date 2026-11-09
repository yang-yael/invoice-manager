#!/bin/zsh
URL="http://127.0.0.1:8765"

if curl -fsS --max-time 1 "$URL" >/dev/null 2>&1; then
  open "$URL"
  exit 0
fi

cd "$(dirname "$0")/invoice_system" || exit 1
python3 web_app.py
