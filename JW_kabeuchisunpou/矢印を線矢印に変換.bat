REM 塗り矢印を線矢印に変換 v2
REM #jww
REM #h1
REM #g1
REM #hc変換したい塗り矢印(三角)を含む範囲を選択（三角ソリッド以外は変更されません）
REM #e
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0arrow_to_line.ps1" -TempFile "jwc_temp.txt"
