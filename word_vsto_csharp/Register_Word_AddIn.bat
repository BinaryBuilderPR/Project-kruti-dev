@echo off
title Register Kruti Dev Word Add-in
echo Registering Kruti Dev 010 Add-in for Microsoft Word...
powershell -ExecutionPolicy Bypass -File "%~dp0Register_Word_AddIn.ps1"
pause
