@echo off
rem Заливка табеля в D365 из терминала.
rem
rem   submit.bat --file data\timesheet.xlsx --week 2026-09-28
rem   submit.bat --file data\timesheet.xlsx            (весь файл)
rem   submit.bat --file data\timesheet.xlsx --preflight-only
rem
rem Почему не `npm start -- --file ...`: в PowerShell npm не доносит флаги до
rem скрипта — до ts-node доезжают только значения ("... src/index.ts
rem data/timesheet.xlsx 2026-09-28"), и commander падает на "required option
rem '-f, --file <path>' not specified". В Git Bash тот же вызов работает, так
rem что баг легко принять за свою ошибку. Здесь node зовётся напрямую — ровно
rem так же, как это делает GUI.
cd /d "%~dp0"
node --require ts-node/register/transpile-only "%~dp0src\index.ts" %*
