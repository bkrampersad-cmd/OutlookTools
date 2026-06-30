@echo off
echo ============================================
echo  Beeran's Outlook Tools v1.5 - Build
echo ============================================
echo.

python --version >nul 2>&1
IF ERRORLEVEL 1 ( echo ERROR: Python not found. & pause & exit /b 1 )

IF NOT EXIST icon.ico (
    echo WARNING: icon.ico not found next to this script.
    echo          The exe will build without a custom icon.
    echo.
)
IF NOT EXIST logo.png (
    echo WARNING: logo.png not found next to this script.
    echo          The sidebar will fall back to plain text.
    echo.
)

echo [1/4] Installing dependencies...
pip install -r requirements.txt
IF ERRORLEVEL 1 ( echo pip install failed. & pause & exit /b 1 )

echo [2/4] pywin32 post-install...
python -c "import win32com" >nul 2>&1
IF ERRORLEVEL 1 (
    python Scripts\pywin32_postinstall.py -install >nul 2>&1
)

echo [3/4] Generating built-in sound files...
python generate_sounds.py
IF ERRORLEVEL 1 ( echo Sound generation failed. & pause & exit /b 1 )

echo [4/4] Building exe...
pyinstaller outlook_tools.spec --clean
IF ERRORLEVEL 1 ( echo Build failed. & pause & exit /b 1 )

echo.
echo ============================================
echo  Done!  dist\BeeransOutlookTools\BeeransOutlookTools.exe
echo ============================================
pause
