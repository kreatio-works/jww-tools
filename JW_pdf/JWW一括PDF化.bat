@echo off
rem Jw_cad drawings -> PDF (one PDF per drawing). Drop .jww files or folders on this file, or double-click.
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0jw_pdf.ps1" %*
