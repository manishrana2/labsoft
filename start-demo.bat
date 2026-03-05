@echo off
setlocal

echo ==========================================
echo Labsoft Demo Launcher (Windows)
echo ==========================================
echo.

where node >nul 2>nul
if %errorlevel% neq 0 (
  echo [ERROR] Node.js not found.
  echo Please install Node.js LTS from https://nodejs.org/
  pause
  exit /b 1
)

echo Installing dependencies (first run may take a while)...
call npm install
if %errorlevel% neq 0 (
  echo [ERROR] npm install failed.
  pause
  exit /b 1
)

echo.
echo Starting Labsoft demo on http://localhost:5173 ...
echo Keep this window open while demo is running.
echo Press Ctrl+C to stop.
echo.
start "" http://localhost:5173
call npm run dev:full
