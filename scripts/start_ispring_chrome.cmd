@echo off
REM Start the Chrome the automation talks to.
REM
REM It uses its own profile folder, so it never fights with your normal
REM browsing, and the debug port lets the scripts drive it. Sign in to
REM iSpring Cloud once in this window; that session is then reused for
REM every publish, with no login again.
REM
REM Leave this Chrome running. On the worker machine, put a shortcut to
REM this file in the Startup folder.

set CHROME="C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist %CHROME% set CHROME="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not exist %CHROME% (
  echo Could not find chrome.exe. Edit this file and set the CHROME path.
  pause
  exit /b 1
)

set PROFILE=C:\ispring-chrome-profile
if not exist "%PROFILE%" mkdir "%PROFILE%"

echo Starting Chrome for iSpring automation...
echo   profile: %PROFILE%
echo   debug port: 9222
echo.
echo Sign in to iSpring Cloud in this window if it asks. Then leave it open.

start "" %CHROME% --remote-debugging-port=9222 --user-data-dir="%PROFILE%" https://harshit.ispring.com/
