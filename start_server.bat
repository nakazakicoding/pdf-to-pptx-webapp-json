@echo off
echo ========================================
echo   PDF to PPTX Converter (JSON Mode)
echo   Port: 8001
echo ========================================
echo.

cd /d "%~dp0"

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    pause
    exit /b 1
)

REM Install dependencies if needed
if not exist "venv" (
    echo [INFO] Creating virtual environment...
    python -m venv venv
)

call venv\Scripts\activate

echo [INFO] Installing dependencies...
pip install -q fastapi uvicorn python-multipart PyMuPDF pillow opencv-python-headless numpy python-pptx pytesseract scikit-image

echo.
echo [INFO] Starting server on http://localhost:8001
echo [INFO] Open your browser to http://localhost:8001/static/index.html
echo.

python server.py

pause
