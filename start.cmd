@echo off
REM Double-click this file (or pin it to the taskbar / Start menu) to launch
REM Math Popup. Uses the bundled Electron binary so it does not need a
REM terminal to be open.
cd /d "%~dp0"
if not exist "node_modules\electron\dist\electron.exe" (
  echo Electron not found. Run: npm install ^&^& npm run build
  pause
  exit /b 1
)
if not exist "dist\main\main.js" (
  echo Build output missing. Running build...
  call npm run build || (pause & exit /b 1)
)
start "" "node_modules\electron\dist\electron.exe" .
