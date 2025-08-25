@echo off
SETLOCAL ENABLEEXTENSIONS
echo -------------------------------
echo Starting full setup...
echo -------------------------------

REM === CHECK IF NODE IS INSTALLED ===
where node >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo Node.js not found. Installing Node.js...

    REM Download Node.js LTS MSI installer
    powershell -Command "Invoke-WebRequest -Uri https://nodejs.org/dist/v18.18.2/node-v18.18.2-x64.msi -OutFile node-lts.msi"

    echo Running Node.js installer...
    msiexec /i node-lts.msi /quiet /norestart

    echo Waiting for Node.js installation to complete...
    timeout /t 15 >nul

    REM Clean up
    del node-lts.msi

    echo Node.js installed successfully.
) ELSE (
    echo Node.js already installed.
)


REM === START SERVER AND CLIENT ===
echo Starting the server and client...
call cmd /k "cd server && npm run build && npm start"

echo -------------------------------
echo All set! React app and server are running.
echo -------------------------------
pause
