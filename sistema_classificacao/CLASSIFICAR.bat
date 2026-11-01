@echo off
chcp 65001 >nul
title Classificacao IBS/CBS/IS
REM Arraste o arquivo do cliente (Excel, CSV, XML ou pasta de XMLs) sobre este arquivo.
if "%~1"=="" (
  echo Arraste o arquivo do cliente sobre o CLASSIFICAR.bat
  pause
  exit /b
)
where python >nul 2>nul || (
  echo Python nao encontrado. Instale em https://www.python.org/downloads/ marcando "Add Python to PATH".
  pause
  exit /b
)
python -c "import openpyxl" 2>nul || python -m pip install --quiet openpyxl
set /p CLIENTE=Nome do cliente: 
set /p REGIME=Regime tributario (Simples / Presumido / Real): 
python "%~dp0classificar.py" "%~1" --cliente "%CLIENTE%" --regime "%REGIME%"
echo.
pause
