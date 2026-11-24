@echo off
rem Start de Intune Management Tool (de console verdwijnt zodra het venster opent)
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0IntuneTool.ps1"
