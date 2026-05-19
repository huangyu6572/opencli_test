@echo off
:: DevCloud task detail extractor
:: Usage:
::   extract.bat <url1> [url2 ...]
::   extract.bat --file urls.txt
::
:: Output: excel\tasks_YYYYMMDD_HHMMSS.csv
setlocal

if "%~1"=="" (
    echo Usage: extract.bat ^<url1^> [url2 ...]
    echo    or: extract.bat --file urls.txt
    exit /b 1
)

set SCRIPT_DIR=%~dp0
python "%SCRIPT_DIR%extract_tasks.py" %*
endlocal
