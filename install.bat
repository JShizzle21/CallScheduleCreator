@echo off
REM ============================================================================
REM Call Schedule Creator — one-time installer
REM ============================================================================
REM What this does:
REM   1. Creates %LOCALAPPDATA%\CallScheduler  (off OneDrive — avoids sync churn
REM      on the ~300MB of pyarrow/pandas binaries we pull in).
REM   2. Downloads Python 3.14.3 embeddable distribution + extracts it.
REM   3. Edits python314._pth to enable site-packages imports.
REM   4. Bootstraps pip via get-pip.py.
REM   5. Installs runtime dependencies from docs\requirements.txt.
REM
REM Idempotent — re-running skips steps that are already done. To force a
REM clean reinstall, delete %LOCALAPPDATA%\CallScheduler first.
REM
REM Run-once: double-click this file. It opens a console window; do not
REM close it until you see "Setup complete." Then double-click run.bat.
REM ============================================================================

setlocal EnableDelayedExpansion

set "PY_VERSION=3.14.3"
set "PY_TAG=314"
set "INSTALL_DIR=%LOCALAPPDATA%\CallScheduler"
set "PY_DIR=%INSTALL_DIR%\python_embed"
set "PY_EXE=%PY_DIR%\python.exe"
set "PY_PTH=%PY_DIR%\python%PY_TAG%._pth"
set "PROJECT_DIR=%~dp0"
set "REQ_FILE=%PROJECT_DIR%docs\requirements.txt"

echo.
echo ============================================================
echo  Call Schedule Creator — Installer
echo ============================================================
echo.
echo  Install location: %INSTALL_DIR%
echo  Project folder:   %PROJECT_DIR%
echo.

if not exist "%REQ_FILE%" (
    echo [ERROR] docs\requirements.txt not found at:
    echo         %REQ_FILE%
    echo.
    echo  This installer must be run from the project folder. Make sure
    echo  the whole project was copied from OneDrive/Teams, not just
    echo  install.bat alone.
    goto :fail
)

REM --- Step 1: Create install directory ----------------------------------------
if not exist "%INSTALL_DIR%" (
    echo [1/5] Creating %INSTALL_DIR% ...
    mkdir "%INSTALL_DIR%" || goto :fail
) else (
    echo [1/5] Install directory already exists — reusing.
)

REM --- Step 2: Download + extract embeddable Python ---------------------------
if exist "%PY_EXE%" (
    echo [2/5] Embeddable Python already present — skipping download.
) else (
    echo [2/5] Downloading Python %PY_VERSION% embeddable ^(~10 MB^) ...
    set "PY_ZIP=%INSTALL_DIR%\python-embed.zip"
    set "PY_URL=https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%-embed-amd64.zip"
    call :download "!PY_URL!" "!PY_ZIP!"
    if errorlevel 1 goto :fail

    echo       Extracting to %PY_DIR% ...
    if not exist "%PY_DIR%" mkdir "%PY_DIR%"
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "Expand-Archive -Path '!PY_ZIP!' -DestinationPath '%PY_DIR%' -Force" || goto :fail
    del "!PY_ZIP!" >nul 2>&1
)

REM --- Step 3: Enable site-packages in ._pth file ------------------------------
if not exist "%PY_PTH%" (
    echo [ERROR] Expected %PY_PTH% not found after extraction.
    goto :fail
)

findstr /B /C:"import site" "%PY_PTH%" >nul 2>&1
if !errorlevel! equ 0 (
    echo [3/5] site-packages already enabled in %PY_TAG%._pth.
) else (
    echo [3/5] Enabling site-packages in python%PY_TAG%._pth ...
    REM Replace "#import site" with "import site"; if no commented line,
    REM append it. Use PowerShell to keep encoding intact.
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "$p='%PY_PTH%'; $t = Get-Content -Raw -LiteralPath $p; if ($t -match '(?m)^\s*#\s*import site') { $t = $t -replace '(?m)^\s*#\s*import site','import site' } else { $t = $t.TrimEnd() + \"`r`nimport site`r`n\" }; Set-Content -LiteralPath $p -Value $t -NoNewline" || goto :fail
)

REM --- Step 4: Bootstrap pip ---------------------------------------------------
"%PY_EXE%" -m pip --version >nul 2>&1
if !errorlevel! equ 0 (
    echo [4/5] pip already installed — skipping bootstrap.
) else (
    echo [4/5] Bootstrapping pip ^(downloading get-pip.py^) ...
    set "PIP_BOOT=%INSTALL_DIR%\get-pip.py"
    call :download "https://bootstrap.pypa.io/get-pip.py" "!PIP_BOOT!"
    if errorlevel 1 goto :fail

    "%PY_EXE%" "!PIP_BOOT!" --no-warn-script-location || goto :fail
    del "!PIP_BOOT!" >nul 2>&1
)

REM --- Step 5: Install runtime requirements -----------------------------------
echo [5/5] Installing runtime dependencies from docs\requirements.txt ...
echo       ^(This is the slow step — ~300 MB of wheels. Be patient.^)
"%PY_EXE%" -m pip install --no-warn-script-location -r "%REQ_FILE%" || goto :fail

echo.
echo ============================================================
echo  Setup complete.
echo  Double-click run.bat to start the Call Schedule Creator.
echo ============================================================
echo.
pause
exit /b 0

REM ----------------------------------------------------------------------------
REM Helper: download URL %1 to file %2. Uses curl (bundled with Win10+); falls
REM back to PowerShell. Returns nonzero on failure.
REM ----------------------------------------------------------------------------
:download
where curl >nul 2>&1
if !errorlevel! equ 0 (
    curl -fSL --retry 3 -o %2 %1
    exit /b !errorlevel!
)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri %1 -OutFile %2 -UseBasicParsing"
exit /b !errorlevel!

:fail
echo.
echo ============================================================
echo  [ERROR] Setup failed. See messages above.
echo ============================================================
echo.
echo  Common causes:
echo    - No internet connection ^(installer needs to download Python + pip^).
echo    - Corporate firewall blocks python.org or pypi.org.
echo    - Antivirus quarantined the downloaded files.
echo    - %INSTALL_DIR% is on a read-only or full disk.
echo.
pause
exit /b 1
