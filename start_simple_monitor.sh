#!/bin/bash

# Simple Excel Claude Integration Monitor
echo "🚀 Starting Simple Excel Claude Integration Monitor..."

# Change to script directory
cd "$(dirname "$0")"

# Check if psutil is installed
python3 -c "import psutil" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "❌ Missing psutil. Installing..."
    python3 -m pip install psutil
fi

# Check if MCP server exists
if [ ! -f "src/excel_mcp/simple_server.py" ]; then
    echo "❌ MCP server not found at src/excel_mcp/simple_server.py"
    exit 1
fi

echo "✅ Starting simple monitor (no GUI, command-line prompts)..."
echo ""

# Start the simple monitor
python3 excel_monitor_simple.py