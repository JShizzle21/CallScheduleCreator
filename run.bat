@echo off
REM ============================================================================
REM Call Schedule Creator — launcher
REM ============================================================================
REM What this does:
REM   1. Cleans up old temp upload dirs (Streamlit doesn't fire a session-end
REM      hook so they accumulate over time).
REM   2. Verifies %LOCALAPPDATA%\CallScheduler\python_embed\python.exe exists
REM      — prompts user to run install.bat if not.
REM   3. Scans for the first free port starting at 8501.
REM   4. Launches Streamlit headless on that port.
REM   5. Opens the user's default browser to that URL.
REM
REM The terminal window stays open while the app runs. Closing it (or pressing
REM Ctrl+C) stops the server.
REM ============================================================================

setlocal EnableDelayedExpansion

set "INSTALL_DIR=%LOCALAPPDATA%\CallScheduler"
set "PY_EXE=%INSTALL_DIR%\python_embed\python.exe"
set "PROJECT_DIR=%~dp0"
set "APP_FILE=%PROJECT_DIR%src\app.py"

REM --- Step 1: Clean up old temp upload directories ---------------------------
REM 7-day-old tmp* dirs from prior sessions. Streamlit's tempfile.mkdtemp dirs
REM have OS-default prefixes ("tmp" on Windows). Quiet on errors.
forfiles /p "%TEMP%" /m "tmp*" /d -7 /c "cmd /c if @isdir==TRUE rmdir /s /q @path" >nul 2>&1

REM --- Step 2: Verify install ran ---------------------------------------------
if not exist "%PY_EXE%" (
    echo.
    echo ============================================================
    echo  [ERROR] Python is not installed.
    echo ============================================================
    echo.
    echo  Expected to find: %PY_EXE%
    echo.
    echo  Please run install.bat first ^(double-click it^), then come
    echo  back and double-click run.bat.
    echo.
    pause
    exit /b 1
)

if not exist "%APP_FILE%" (
    echo [ERROR] app.py not found at %APP_FILE%
    echo  This launcher must live in the project folder alongside the src\ subfolder.
    pause
    exit /b 1
)

REM --- Step 3: Find a free port (8501..8520) ----------------------------------
set "PORT="
for /l %%P in (8501,1,8520) do (
    if not defined PORT (
        netstat -an | findstr /R /C:":%%P .*LISTENING" >nul 2>&1
        if errorlevel 1 set "PORT=%%P"
    )
)
if not defined PORT (
    echo [ERROR] All ports 8501-8520 are in use. Close some apps and retry.
    pause
    exit /b 1
)

REM --- Step 4: Launch -----------------------------------------------------------
echo.
echo ============================================================
echo  Call Schedule Creator
echo ============================================================
echo.
echo  Starting server on http://localhost:!PORT!  ...
echo  ^(First launch takes 5-15 seconds. The browser will open
echo   automatically. Keep this window open while you use the app.^)
echo.

REM Streamlit's --server.headless=true means it does NOT auto-open a browser.
REM We open it explicitly after a brief delay so the server has time to bind.
start "" /b cmd /c "timeout /t 4 /nobreak >nul && start "" http://localhost:!PORT!"

cd /d "%PROJECT_DIR%"
"%PY_EXE%" -m streamlit run "%APP_FILE%" ^
    --server.port=!PORT! ^
    --server.headless=true ^
    --browser.gatherUsageStats=false ^
    --server.fileWatcherType=none

echo.
echo Server stopped.
pause
endlocal
