@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Para usar PIN, troque a linha abaixo por:  python servidor.py --pin 1234
python servidor.py
pause
