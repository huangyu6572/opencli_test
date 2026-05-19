@echo off
:: DevCloud task detail URLs -> demo.xlsx
:: Usage: update_demo.bat <url1> [url2 ...] OR update_demo.bat --file urls.txt
setlocal
if "%~1"=="" (
    echo Usage: update_demo.bat ^<url1^> [url2 ...]
    echo    or: update_demo.bat --file urls.txt
    exit /b 1
)
set SCRIPT_DIR=%~dp0
python "%SCRIPT_DIR%update_demo.py" %*
endlocal
