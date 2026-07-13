@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0TOOL_Check_CCDC\Chay_Check_CCDC.ps1" -Root "%~dp0."
pause
