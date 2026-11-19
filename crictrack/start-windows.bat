@echo off
REM Double-click to start cricpay. Set or change the edit PIN from the website sidebar ("Set edit PIN").
cd /d "%~dp0"
python -m pip install -r requirements.txt --quiet
python app.py
pause
