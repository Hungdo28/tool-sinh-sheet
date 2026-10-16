@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process (Join-Path (Get-Location).Path 'web\index.html')"
