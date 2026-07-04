::[Bat To Exe Converter]
::
::YAwzoRdxOk+EWAjk
::fBw5plQjdCyDJGyX8VAjFDpXWwyBAE+/Fb4I5/jH3+OLp0AcRNItd4Xe2aCyGeEB7kjlZaok1XVUi/cEDQlLfR3lZww7yQ==
::YAwzuBVtJxjWCl3EqQJgSA==
::ZR4luwNxJguZRRnk
::Yhs/ulQjdF+5
::cxAkpRVqdFKZSDk=
::cBs/ulQjdF+5
::ZR41oxFsdFKZSDk=
::eBoioBt6dFKZSDk=
::cRo6pxp7LAbNWATEpCI=
::egkzugNsPRvcWATEpCI=
::dAsiuh18IRvcCxnZtBJQ
::cRYluBh/LU+EWAnk
::YxY4rhs+aU+JeA==
::cxY6rQJ7JhzQF1fEqQJQ
::ZQ05rAF9IBncCkqN+0xwdVs0
::ZQ05rAF9IAHYFVzEqQJQ
::eg0/rx1wNQPfEVWB+kM9LVsJDGQ=
::fBEirQZwNQPfEVWB+kM9LVsJDGQ=
::cRolqwZ3JBvQF1fEqQJQ
::dhA7uBVwLU+EWDk=
::YQ03rBFzNR3SWATElA==
::dhAmsQZ3MwfNWATElA==
::ZQ0/vhVqMQ3MEVWAtB9wSA==
::Zg8zqx1/OA3MEVWAtB9wSA==
::dhA7pRFwIByZRRnk
::Zh4grVQjdCyDJGyX8VAjFDpXWwyBAE+/Fb4I5/jH3+OLp0AcRNItd4Xe2aCyGeEB7kjlZaoU12helcocOB5LalyudgpU
::YB416Ek+ZG8=
::
::
::978f952a14a936cc963da21a135fa983
@echo off
title Yardmap Server

REM Check if port 8765 is already in use (prior session still running)
netstat -ano | findstr ":8765 " | findstr "LISTENING" > nul
if %ERRORLEVEL% EQU 0 (
    echo Yardmap is already running. Opening browser...
    start http://localhost:8765
    exit /b
)

REM Open the browser after a short delay (side process so this window stays for the server)
start "" cmd /c "timeout /t 2 /nobreak > nul & start http://localhost:8765"

REM Run the server directly in this window -- close this window to stop
cd /d "%~dp0"
echo Yardmap Server running at http://localhost:8765
echo Close this window to stop the server.
echo.
python -m http.server 8765
