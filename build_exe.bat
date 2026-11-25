@echo off
REM Builds two executables in dist\ (run on Windows with Python installed):
REM   invoice_checks_web.exe  -> opens the web interface in the browser
REM   invoice_checks.exe      -> command line / menu
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller
pyinstaller --onefile --console --name invoice_checks_web web.py
pyinstaller --onefile --console --name invoice_checks invoice_checks.py
copy config.json dist\
xcopy /E /I /Y data dist\data
xcopy /E /I /Y web dist\web
echo.
echo Done: dist\invoice_checks_web.exe
pause
