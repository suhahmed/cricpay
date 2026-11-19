#!/usr/bin/env bash
# Start cricpay AND share it on your fixed public ngrok address, in one go.
# One-time setup: see "Same public link every time" in README.md (ngrok account + authtoken).
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt --quiet

NGROK="$(command -v ngrok || true)"
[ -z "$NGROK" ] && [ -x "$HOME/Downloads/ngrok" ] && NGROK="$HOME/Downloads/ngrok"
if [ -z "$NGROK" ]; then
  echo "ngrok not found. Put the ngrok file in your Downloads folder (see README.md)."; exit 1
fi

python3 app.py & APP=$!
caffeinate -dimsu -w $APP 2>/dev/null &          # keep the Mac awake while cricpay runs
trap 'kill $APP 2>/dev/null' EXIT INT TERM       # Ctrl+C stops both
sleep 4
echo ""
echo "  Your public cricpay link is the https://....ngrok-free.app address shown below."
echo "  Press Ctrl+C to stop."
echo ""
"$NGROK" http 8000
