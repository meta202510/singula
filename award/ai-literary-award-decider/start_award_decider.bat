@echo off
cd /d "%~dp0"
py -3 award_decider_gui.py
if errorlevel 1 pause
