@echo off
:: DevCloud Sprint Tasks → Excel one-liner
:: Usage: export.bat "https://hn.devcloud.huaweicloud.com/.../sprint/<id>/list"
:: Output: excel\sprint_tasks_<timestamp>.xlsx
setlocal

if "%~1"=="" (
    echo Usage: export.bat ^<sprint-url^>
    echo Example: export.bat "https://hn.devcloud.huaweicloud.com/projectman/scrum/3ae70814de3545d3a3d90e3d1dd4bbb7/task/sprint/721752983/list"
    exit /b 1
)

set SPRINT_URL=%~1
set SCRIPT_DIR=%~dp0

echo Fetching sprint tasks from DevCloud...
opencli devcloud sprint-tasks --url "%SPRINT_URL%" --format csv | python "%SCRIPT_DIR%excel\export_to_excel.py"
if errorlevel 1 (
    echo ERROR: Export failed. Make sure Chrome is open and logged into DevCloud.
    exit /b 1
)
endlocal
