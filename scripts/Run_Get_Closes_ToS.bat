@echo off
setlocal
cd /d "C:\Users\ryanc\OneDrive\repos\Trading_Algo\scripts"

set "PY=C:\Users\ryanc\AppData\Local\Python\pythoncore-3.14-64\python.exe"
set "LOG=C:\Users\ryanc\OneDrive\repos\Trading_Algo\Admin\Logs\Get_Closes_ToS.log"
set "EARNINGS=C:\Users\ryanc\OneDrive\repos\Trading_Algo\! -- Latest Earnings Document.xlsx"

if not exist "C:\Users\ryanc\OneDrive\repos\Trading_Algo\Admin\Logs" mkdir "C:\Users\ryanc\OneDrive\repos\Trading_Algo\Admin\Logs"

echo ===== %date% %time% bat wrapper ===== >> "%LOG%"

REM Close Excel and clear read-only so the script can open/save the workbook.
REM stderr discarded so "EXCEL.EXE not found" is not logged when Excel is already closed.
taskkill /IM EXCEL.EXE /F /T >> "%LOG%" 2>nul
attrib -R "%EARNINGS%" >> "%LOG%" 2>&1
timeout /t 10 /nobreak >nul

echo Running Get_Closes_ToS.py (script tees its own log) >> "%LOG%"
set "ALGO_SCHEDULED=1"
REM Do not redirect stdout here — the .py appends to the same log via run_logging.
"%PY%" "Get_Closes_ToS.py"
set "EXIT_CODE=%ERRORLEVEL%"
echo Exit code: %EXIT_CODE% >> "%LOG%"

REM Cleanup any orphan hidden Excel left from a prior run.
taskkill /IM EXCEL.EXE /F /T >> "%LOG%" 2>nul
exit /b %EXIT_CODE%