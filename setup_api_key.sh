#!/usr/bin/env bash
# Setup helper: create .env and remind how to add a free Gemini API key.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

if [[ ! -f .env ]]; then
  cp .env.example .env
  echo "Created .env from .env.example"
fi

echo ""
echo "=== Free vision API (Gemini) ==="
echo "1. Open:  https://aistudio.google.com/apikey"
echo "2. Click Create API key (Google account, free tier)"
echo "3. Paste into .env as:  GEMINI_API_KEY=your_key_here"
echo "4. Restart:  uvicorn app:app --host 0.0.0.0 --port 8000"
echo ""
echo "OCR works WITHOUT a key. Vision assist turns on when GEMINI_API_KEY is set."
echo ""

KEY_LINE="$(grep -E '^GEMINI_API_KEY=' .env | head -1 || true)"
KEY="${KEY_LINE#GEMINI_API_KEY=}"
KEY="${KEY%\"}"
KEY="${KEY#\"}"
KEY="${KEY%\'}"
KEY="${KEY#\'}"
KEY="$(echo "$KEY" | tr -d '[:space:]')"

if [[ -n "$KEY" && "$KEY" != "your_key_here" ]]; then
  echo "GEMINI_API_KEY is set in .env (length ${#KEY})."
else
  echo "GEMINI_API_KEY is still empty — paste your free key into .env"
fi
