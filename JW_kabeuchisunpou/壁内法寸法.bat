REM 壁内法寸法 v4
REM #jww
REM #h1
REM #hc内法を計測する壁面線を範囲選択（2本=そのまま計測 / 複数=最小間隔で自動抽出）
REM #e
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0wall_clearance.ps1" -TempFile "jwc_temp.txt" -MinGap 300 -AngleTol 1 -Offset 0 -RoundStep 1 -Orient para -Term line -Font 3
