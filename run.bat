@echo off
cd /d "%~dp0"
powershell -Command "python -u run_nav_analysis.py 2>&1 | Tee-Object -FilePath nav_log.txt -Append"
