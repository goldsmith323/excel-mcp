@echo off
:: Simple Excel Claude Integration Monitor for Windows
echo Starting Simple Excel Claude Integration Monitor...

:: Change to script directory
cd /d %~dp0

:: Check if psutil is installed
python -c "import psutil" >nul 2>&1
if errorlevel 1 (
    echo Missing psutil. Installing...
    python -m pip install psutil
)

:: Check if MCP server exists
if not exist "src\excel_mcp\simple_server.py" (
    echo MCP server not found at src\excel_mcp\simple_server.py
    exit /b 1
)

echo Starting simple monitor (no GUI, command-line prompts)...
echo.

:: Start the simple monitor
python excel_monitor_simple.py