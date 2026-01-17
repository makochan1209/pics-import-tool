@echo off
REM バッチファイル自身があるディレクトリを取得
set "SCRIPT_DIR=%~dp0"

REM カレントディレクトリをそのディレクトリに変更
cd /d "%SCRIPT_DIR%"

REM バッチファイルと同名の .ps1 を実行（引数もそのまま渡す）
set "PS1_FILE=%~n0.ps1"

REM ★ここを powershell.exe ではなく pwsh.exe にする★
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%%PS1_FILE%" %*
