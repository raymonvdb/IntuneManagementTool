@echo off
rem Starts the Intune Management Tool (the console disappears as soon as the window opens)
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0IntuneTool.ps1"
