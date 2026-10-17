@echo off
REM Gera o executavel AutomacaoCOAF.exe (rodar UMA vez, em um PC Windows com Python 3 e Excel)
set PY=py
where py >nul 2>nul || set PY=python
where %PY% >nul 2>nul || (
  echo Python nao encontrado. Instale em python.org/downloads marcando "Add python.exe to PATH" e rode de novo.
  pause
  exit /b 1
)
if not exist config.ini (
  echo config.ini nao encontrado. Copie config.example.ini para config.ini e ajuste os valores.
  pause
  exit /b 1
)
if not exist modelo\modelo_planilha.xls (
  echo Coloque a sua planilha modelo em modelo\modelo_planilha.xls e rode de novo.
  pause
  exit /b 1
)
%PY% -m pip install --upgrade pyinstaller -r requirements.txt
%PY% -m PyInstaller --onefile --name AutomacaoCOAF --collect-all pdfplumber --collect-all pdfminer --hidden-import win32timezone --hidden-import pyodbc coaf_automacao.py
if not exist dist\modelo mkdir dist\modelo
copy /Y config.ini dist\config.ini
copy /Y modelo\modelo_planilha.xls dist\modelo\modelo_planilha.xls
echo.
echo Pronto: use a pasta "dist" (AutomacaoCOAF.exe + config.ini + modelo).
pause
