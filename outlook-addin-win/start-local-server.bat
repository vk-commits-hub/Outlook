@echo off
:: ============================================================
::  AI Reply — Local HTTPS server for testing
::  (Only needed if you're NOT using GitHub Pages)
:: ============================================================

title AI Reply — Local Server

echo.
echo   AI Reply Assistant — Local Test Server
echo   ----------------------------------------
echo.

:: Check Node.js
where node >nul 2>&1
if %errorlevel% neq 0 (
    echo   ERROR: Node.js is not installed.
    echo   Download it from https://nodejs.org and re-run this.
    pause
    exit /b 1
)

:: Install tools if missing
call npx --yes office-addin-dev-certs install 2>nul
if %errorlevel% neq 0 (
    echo   Note: Could not auto-install dev certificate.
    echo   You may need to run this as Administrator once.
)

:: Get cert paths
for /f "delims=" %%i in ('npx office-addin-dev-certs get-cert 2^>nul') do set CERT=%%i
for /f "delims=" %%i in ('npx office-addin-dev-certs get-key  2^>nul') do set KEY=%%i

:: Start server
echo   Starting HTTPS server at https://localhost:3000
echo   Leave this window open while using the add-in.
echo   Press Ctrl+C to stop.
echo.

:: Update manifest.xml to use localhost (only if still has placeholder)
powershell -Command "(Get-Content manifest.xml) -replace 'YOUR_GITHUB_PAGES_URL','https://localhost:3000' | Set-Content manifest.xml" 2>nul

npx http-server . -p 3000 --ssl --cert "%CERT%" --key "%KEY%" --cors -c-1

pause
