@echo off
SET scriptPath=%~dp0
PowerShell.exe -ExecutionPolicy Bypass -File "%scriptPath%CheckSuspiciousActivity.ps1"
pause
