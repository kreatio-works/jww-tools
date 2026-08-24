REM 図面差分チェッカー 旧図保存 v1
REM #jww
REM #g1
REM #h1
REM #hc差分対象を全選択してください（旧図面をスナップショット保存）
REM #e
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0jw_diff.ps1" -Mode save -TempFile "jwc_temp.txt" -BaseFile "jw_diff_base.txt"
