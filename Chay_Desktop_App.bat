@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command "python main.py"
if %errorlevel% neq 0 pause
