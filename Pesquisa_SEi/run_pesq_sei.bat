@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set "PYTHON=C:\Users\felipe.rsouza\Automação SNEAELIS\Felipe-\.venv\Scripts\python.exe"
set "SCRIPT=C:\Users\felipe.rsouza\Automação SNEAELIS\Felipe-\Pesquisa_SEi\Pesquisa_SEi_1-2.py"

if not exist "%PYTHON%" (
    echo Python not found: "%PYTHON%"
    pause
    exit /b 1
)

if not exist "%SCRIPT%" (
    echo Script not found: "%SCRIPT%"
    pause
    exit /b 1
)

set "PORT1=9222"
set "PORT2=9224"
set "PORT3=9226"
set "PORT4=9228"

echo Opening 4 terminals...
echo.

for %%Q in (1 2 3 4) do (
    call set "PORT=%%PORT%%Q%%"
    
    REM Open terminal with title showing quarter and port
    start "Quarter %%Q - Port !PORT!" cmd /k "title Quarter %%Q - Port !PORT! && echo ===================================================== && echo Running Quarter %%Q on port !PORT! && echo ===================================================== && echo. && echo When prompted, enter: && echo   Quarter: %%Q && echo   Port: !PORT! && echo. && "%PYTHON%" "%SCRIPT%""
    
    REM Wait 2 seconds between opening windows
    timeout /t 2 /nobreak >nul
)

echo.
echo All 4 terminals opened!
echo.
echo For each terminal, when asked:
echo   - Digite qual quarto: enter %%Q
echo   - Enter the door: enter the corresponding port
echo.
pause