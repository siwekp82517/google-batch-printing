@echo off
echo Starting Google Batch PDF Printer...
echo.
echo Opening http://localhost:8080
echo Press Ctrl+C to stop the server.
echo.

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    start http://localhost:8080
    python -m http.server 8080
    goto :eof
)

where python3 >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    start http://localhost:8080
    python3 -m http.server 8080
    goto :eof
)

echo ERROR: Python not found.
echo Please install Python from https://www.python.org/
echo Or run start-server.ps1 which has a .NET fallback.
pause
