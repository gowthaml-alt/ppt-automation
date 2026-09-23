@echo off
REM ===========================================================
REM  ppt-automation  -  the one command the support team needs
REM
REM    ppt-automation run       keep publishing decks until you stop it
REM    ppt-automation browser   start the Chrome it signs in to
REM    ppt-automation health    is this machine allowed to take jobs?
REM    ppt-automation setup     first time on a new machine
REM
REM  It finds its own folder, turns the Python environment on, and
REM  checks the things that people forget. Nothing to remember, and
REM  no need to be in the right directory.
REM ===========================================================
setlocal

REM Work from the folder this file lives in, whatever the user's
REM current directory is. .env is read from there, so this matters.
pushd "%~dp0"

set ACTION=%~1
if "%ACTION%"=="" set ACTION=run

if /i "%ACTION%"=="setup"   goto do_setup
if /i "%ACTION%"=="run"     goto need_env
if /i "%ACTION%"=="health"  goto need_env
if /i "%ACTION%"=="browser" goto do_browser
goto do_help

REM -----------------------------------------------------------
:need_env
if not exist ".venv\Scripts\activate.bat" (
  echo.
  echo   Python is not set up on this machine yet.
  echo   Run this first:  ppt-automation setup
  echo.
  goto stop_fail
)
call ".venv\Scripts\activate.bat"

if not exist ".env" (
  echo.
  echo   There is no .env file, so the worker does not know which
  echo   server to talk to. Copying the example one for you.
  echo.
  copy /y ".env.example" ".env" >nul
  echo   Open .env and check BACKEND_BASE_URL before running again.
  echo.
  goto stop_fail
)

if /i "%ACTION%"=="health" goto do_health
goto do_run

REM -----------------------------------------------------------
:do_run
echo.
echo   Starting the PPT worker. Press Ctrl+C to stop it.
echo.
python run.py
goto stop_ok

REM -----------------------------------------------------------
:do_health
python scripts\worker_health.py %2 %3
goto stop_ok

REM -----------------------------------------------------------
:do_browser
call scripts\start_ispring_chrome.cmd
goto stop_ok

REM -----------------------------------------------------------
:do_setup
echo.
echo   Setting up Python for the PPT worker. This takes a few minutes.
echo.
if not exist ".venv\Scripts\activate.bat" (
  python -m venv .venv
  if errorlevel 1 (
    echo.
    echo   Could not create the Python environment. Is Python installed?
    echo   Get it from python.org and tick "Add python.exe to PATH".
    echo.
    goto stop_fail
  )
)
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip >nul
pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo   Could not install what the worker needs. Check the internet
  echo   connection and run  ppt-automation setup  again.
  echo.
  goto stop_fail
)
if not exist ".env" copy /y ".env.example" ".env" >nul
echo.
echo   Done. Two things before the first run:
echo.
echo     1. Open .env and check BACKEND_BASE_URL is the right server.
echo     2. Run  ppt-automation browser  and sign in to iSpring Cloud once.
echo.
echo   Then:  ppt-automation run
echo.
goto stop_ok

REM -----------------------------------------------------------
:do_help
echo.
echo   ppt-automation run       publish decks until you stop it (Ctrl+C)
echo   ppt-automation browser   start the Chrome it signs in to
echo   ppt-automation health    is this machine allowed to take jobs?
echo   ppt-automation setup     first time on a new machine
echo.
goto stop_ok

REM -----------------------------------------------------------
:stop_fail
popd
endlocal
exit /b 1

:stop_ok
popd
endlocal
exit /b 0
