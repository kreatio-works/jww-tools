REM 図面差分チェッカー 比較 v1
REM #jww
REM #g1
REM #h1
REM #hc差分対象を全選択してください（新図面を旧図と比較してマーキング）
REM #e
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0jw_diff.ps1" -Mode compare -TempFile "jwc_temp.txt" -BaseFile "jw_diff_base.txt" -Tol 0.5 -MarkR 100 -ColorDel 4 -ColorAdd 5 -Layer F -MaxMark 500
