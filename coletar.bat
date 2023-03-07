@echo off
call "%~dp0Python\scripts\env_for_icons.bat"  %*
python coletar.py
pause
