@echo off
echo Installing required libraries (OpenThai OCR runtime, etc.). This only takes a minute...
pip install -r "%~dp0requirements.txt"
echo.
echo Starting Unit Test Image Injector (Python Script Mode)...
python "%~dp0app.py"
pause
