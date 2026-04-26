@echo off
REM ============================================================================
REM Call Schedule Creator — uninstaller
REM ============================================================================
REM Removes everything install.bat created:
REM   - %LOCALAPPDATA%\CallScheduler\  (embedded Python + dependencies, ~400 MB)
REM   - %TEMP%\CallScheduler_*  (any leftover upload staging dirs)
REM
REM Does NOT touch the project folder itself or your data files. Delete the
REM project folder by hand from Explorer when you are done.
REM ============================================================================

setlocal EnableDelayedExpansion

set "INSTALL_DIR=%LOCALAPPDATA%\CallScheduler"

echo.
echo ============================================================
echo  Call Schedule Creator — Uninstaller
echo ============================================================
echo.

if not exist "%INSTALL_DIR%" (
    echo  Nothing to uninstall — %INSTALL_DIR% does not exist.
    echo.
    echo  ^(If you also want to remove the program files, delete this
    echo   folder manually from Windows Explorer.^)
    echo.
    pause
    exit /b 0
)

echo  This will permanently delete:
echo    %INSTALL_DIR%
echo    ^(embedded Python and all dependencies, about 400 MB^)
echo.
echo  It will NOT delete your project folder or your data files.
echo  Those have to be removed by hand if you want them gone too.
echo.

set /p CONFIRM="  Continue? Type Y to uninstall, anything else to cancel: "
if /i not "!CONFIRM!"=="Y" (
    echo.
    echo  Uninstall cancelled. No changes made.
    echo.
    pause
    exit /b 0
)

echo.
echo  Removing %INSTALL_DIR% ...
rmdir /s /q "%INSTALL_DIR%"
if exist "%INSTALL_DIR%" (
    echo.
    echo  [ERROR] Could not fully remove %INSTALL_DIR%.
    echo  Possible causes:
    echo    - The app is still running. Close it ^(close the run.bat
    echo      window or click "Exit and shut down" in the browser^),
    echo      then re-run this uninstaller.
    echo    - A file is open in another program.
    echo    - You do not have permission to delete the folder.
    echo.
    pause
    exit /b 1
)

REM Best-effort cleanup of any leftover upload staging dirs from old runs.
echo  Cleaning up temporary upload directories ...
for /d %%D in ("%TEMP%\CallScheduler_*") do rmdir /s /q "%%D" >nul 2>&1

echo.
echo ============================================================
echo  Uninstall complete.
echo ============================================================
echo.
echo  To remove the program files too, delete this folder from
echo  Windows Explorer ^(it has install.bat, run.bat, README.md,
echo   and the data\ subfolder in it^).
echo.
echo  To reinstall later, double-click install.bat.
echo.
pause
exit /b 0
