@echo off
setlocal
cd /d "%~dp0"

if not exist ".env" (
    echo ERROR: File .env was not found in %CD%
    echo Copy .env.example to .env and fill in the mailbox settings.
    exit /b 1
)

set "AGENT_PYTHON=%CD%\.venv\Scripts\python.exe"

if not exist "%AGENT_PYTHON%" (
    echo Creating Python virtual environment...
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3.11 -m venv .venv
    ) else (
        python -m venv .venv
    )
    if errorlevel 1 (
        echo ERROR: Failed to create .venv. Install Python 3.11 or newer.
        exit /b 1
    )
)

"%AGENT_PYTHON%" -c "import openpyxl, pydantic, pydantic_settings" >nul 2>nul
if errorlevel 1 (
    echo Installing project dependencies...
    "%AGENT_PYTHON%" -m pip install -e .
    if errorlevel 1 (
        echo ERROR: Failed to install project dependencies.
        exit /b 1
    )
)

set "PYTHONPATH=%CD%\src"

if "%~1"=="" (
    echo Specify one agent to run once:
    echo   .\run-agent.cmd request-agent
    echo   .\run-agent.cmd result-agent
    echo.
    echo Use --continuous for a polling loop, or run-agents for both loops.
    "%AGENT_PYTHON%" -m shipment_forecast_agent.cli --help
) else (
    "%AGENT_PYTHON%" -m shipment_forecast_agent.cli %*
)

exit /b %ERRORLEVEL%
