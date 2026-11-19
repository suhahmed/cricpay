#!/usr/bin/env bash
# Start cricpay. Set or change the edit PIN from the website sidebar ("Set edit PIN").
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt --quiet
python3 app.py
