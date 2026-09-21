@echo off
echo ============================================================
echo CONSENT FORM COMPLETE WORKFLOW
echo ============================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

REM Check if virtual environment exists
if not exist "venv\Scripts\activate.bat" (
    echo Virtual environment not found. Creating it now...
    echo.
    echo ============================================================
    echo SETTING UP VIRTUAL ENVIRONMENT
    echo ============================================================
    echo.
    
    echo Creating virtual environment...
    python -m venv venv
    
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment
        echo.
        pause
        exit /b 1
    )
    
    echo.
    echo Activating virtual environment...
    call venv\Scripts\activate.bat
    
    echo.
    echo Upgrading pip...
    python -m pip install --upgrade pip
    
    echo.
    echo Installing required packages...
    python -m pip install -r requirements.txt
    
    if errorlevel 1 (
        echo ERROR: Failed to install required packages
        echo.
        pause
        exit /b 1
    )
    
    echo.
    echo ============================================================
    echo SETUP COMPLETE!
    echo ============================================================
    echo.
) else (
    echo Virtual environment found. Activating...
    call venv\Scripts\activate.bat
    
    REM Check if packages are installed by testing pandas import
    python -c "import pandas" 2>nul
    if errorlevel 1 (
        echo Packages not found. Installing required packages...
        echo.
        python -m pip install --upgrade pip
        python -m pip install -r requirements.txt
        echo.
        echo Packages installed successfully!
        echo.
    )
)

echo.
echo Starting the workflow...
echo.
echo ============================================================
echo.

REM Run the Python script
python consent_form_complete_workflow.py

echo.
echo ============================================================
echo WORKFLOW FINISHED
echo ============================================================
echo.
pause
