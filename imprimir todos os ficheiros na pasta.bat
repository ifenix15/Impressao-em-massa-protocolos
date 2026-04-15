@echo off
setlocal enabledelayedexpansion

REM -------------------------------------------
REM CAMINHO DO ACROBAT
REM -------------------------------------------
set "ACROBAT=C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

if not exist "%ACROBAT%" (
    echo ERRO: O Acrobat Pro nao foi encontrado no caminho:
    echo %ACROBAT%
    pause
    exit /b
)

REM -------------------------------------------
REM LISTAR IMPRESSORAS
REM -------------------------------------------
echo A obter lista de impressoras...
set /a COUNT=0

for /f "tokens=* usebackq" %%P in (`
    powershell -NoProfile -Command "Get-Printer | Select-Object -ExpandProperty Name"
`) do (
    set /a COUNT+=1
    set "PRN!COUNT!=%%P"
)

echo Impressoras disponiveis:
for /l %%I in (1,1,%COUNT%) do (
    echo [%%I] !PRN%%I!
)

echo.
set /p CHOICE="Escolha a impressora (numero): "
set "PRINTER=!PRN%CHOICE%!"

echo.
echo Impressora escolhida: %PRINTER%
echo.

REM -------------------------------------------
REM ABRIR TODOS OS FICHEIROS PERMITIDOS
REM -------------------------------------------
echo A abrir ficheiros no Acrobat...

for %%F in (*.pdf *.doc *.docx) do (
    echo A abrir: %%F
    start "" "%ACROBAT%" "%%~fF"
)

REM Pequena pausa para garantir que todos abriram
timeout /t 3 >nul

REM -------------------------------------------
REM IMPRIMIR CADA FICHEIRO INDIVIDUALMENTE
REM -------------------------------------------
echo.
echo A imprimir todos os ficheiros...

for %%F in (*.pdf) do (
    echo A imprimir PDF: %%F
    "%ACROBAT%" /t "%%~fF" "%PRINTER%"
)

for %%F in (*.doc *.docx) do (
    echo A imprimir Word: %%F
    powershell -Command ^
    "$a = New-Object -ComObject Shell.Application; $a.ShellExecute('%%~fF', '', '', 'printto', '%PRINTER%')"
)

echo.
echo ✅ Todos os ficheiros foram enviados para a impressora: %PRINTER%
echo.
pause
