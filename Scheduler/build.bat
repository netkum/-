@echo off
rem Team scheduler builder - run from this folder
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"
pause
