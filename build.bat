@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

if "%~1"=="--clean" (
  if exist build rmdir /s /q build
  if exist dist  rmdir /s /q dist
)

C:\Python314\python.exe -m PyInstaller --noconfirm --clean main.spec
set ERR=%ERRORLEVEL%
if not %ERR%==0 (
  echo Erro ao compilar. Codigo: %ERR%
  exit /b %ERR%
)

echo.
if exist dist\main.exe (
  echo EXE gerado em: dist\main.exe
) else if exist dist\main\main.exe (
  echo EXE gerado em: dist\main\main.exe
) else (
  echo Concluido, mas nao encontrei o main.exe em dist\. Verifique a saida acima.
)

endlocal
