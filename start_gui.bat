@echo off
cd /d "%~dp0"
py -3 novel_critique_gui.py
if errorlevel 1 pause
