@echo off
REM ===============================================================
REM  BENZAITEN - rotina diaria com dois cliques
REM  Sorteia a escala de hoje e gera o relatorio da URA.
REM ===============================================================
chcp 65001 > nul
title BENZAITEN
cd /d "%~dp0"

REM Usa "python"; se nao existir, tenta o "py" (instalador do Windows)
set "PY=python"
where python >nul 2>nul || set "PY=py"

echo.
echo   ================================================
echo     BENZAITEN - rotina diaria
echo   ================================================
echo.
echo   Antes de continuar, confira se estao na pasta "entrada":
echo     - a esteira do Vanguard
echo     - as tarifacoes VENDITORE e 4NET
echo.
set "AUSENTES="
set /p "AUSENTES=  Ausentes na escala de hoje (separe por virgula; Enter se ninguem): "

if defined AUSENTES (
    %PY% benzaiten.py --ausentes "%AUSENTES%"
) else (
    %PY% benzaiten.py
)

echo.
pause
