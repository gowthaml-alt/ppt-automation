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
setlocal enabledelayedexpansion

REM Work from the folder this file lives in, whatever the user's
REM current directory is. .env is read from there, so this matters.
pushd "%~dp0"

REM Was this double-clicked in Explorer, or typed at a prompt?
REM Explorer runs it with /c, which closes the window the moment the
REM script ends - taking any error message with it. When that is how we
REM were started, wait for a key at the end so it can be read.
set KEEPOPEN=
echo %cmdcmdline% | find /i "%~nx0" >nul
if not errorlevel 1 set KEEPOPEN=1

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
REM Check the machine before starting, and say so plainly.
REM
REM Three failures in a row that point at this machine and the worker
REM stops taking jobs, so one broken machine cannot mark the whole
REM queue failed. It stays stopped until somebody says otherwise -
REM clearing it automatically here would undo the only thing that
REM protects the queue. So it is reported, and the answer is a person's.
python scripts\worker_health.py
if errorlevel 1 (
  echo.
  echo   This machine is marked UNHEALTHY and will not take any jobs.
  echo.
  echo   The failures above are why. If the machine has been put right -
  echo   iSpring signed in, PowerPoint working - it can go back in service.
  echo.
  set /p BACKINSERVICE="   Put it back in service and start? (y/N): "
  if /i not "!BACKINSERVICE!"=="y" (
    echo.
    echo   Left as it is. Nothing was started.
    echo.
    goto stop_fail
  )
  python scripts\worker_health.py --clear
)

echo.
echo   Starting the PPT worker. Press Ctrl+C to stop it.
echo   It opens Chrome and PowerPoint by itself when it needs them.
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
if defined KEEPOPEN pause
popd
endlocal
exit /b 1

:stop_ok
if defined KEEPOPEN pause
popd
endlocal
exit /b 0
