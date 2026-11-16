@echo off
rem Analysis Tracker launcher (no Excel needed). Keep this file next to the AnalizTakip folder.
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0AnalizTakip\AnalizTakip.ps1"
