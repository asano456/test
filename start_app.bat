@echo off
cd /d %~dp0

REM Flask統合アプリを起動（1つのターミナルでOK）
start cmd /k "python flask_mail_full_app.py"

REM アプリの起動を待つ（Flaskが準備完了するまで）
timeout /t 5 >nul

REM HTMLページ（ルート）をブラウザで開く
start "" http://127.0.0.1:5000/
start "" http://127.0.0.1:5000/log-view
start "" http://127.0.0.1:5000/log-history
start "" http://127.0.0.1:5000/analyze
