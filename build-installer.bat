@echo off
REM One-click installer build (double-click entry). Forwards to build-installer.ps1.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-installer.ps1" %*
pause
