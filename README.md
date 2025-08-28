# Excel MCP Server

A Model Context Protocol (MCP) server for Excel file operations using Python.

**Author:** Hossein Zargar

## Features

- **get_document_info**: Get basic information about Excel documents including sheet names, dimensions, and metadata
- **update_cell**: Update cell values using A1 notation (e.g., "A1", "B5")

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the server:

**On Linux/macOS:**
```bash
python3 -m src.excel_mcp.simple_server
```

**On Windows:**
```cmd
python -m src.excel_mcp.simple_server
```

## Usage

Set the `EXCEL_FILE_PATH` environment variable to specify your Excel file:

**On Linux/macOS:**
```bash
EXCEL_FILE_PATH=/path/to/your/file.xlsx python3 -m src.excel_mcp.simple_server
```

**On Windows:**
```cmd
set EXCEL_FILE_PATH=path\to\your\file.xlsx
python -m src.excel_mcp.simple_server
```

## Quick Start

The repository includes sample Excel files for testing. To get started quickly:

**On Linux/macOS:**
```bash
./start_simple_monitor.sh
```

**On Windows:**
```cmd
start_simple_monitor.bat
```

## MCP Tools

### get_document_info
Returns information about the Excel file including:
- File path and size
- Number of sheets
- Sheet names and dimensions

### update_cell
Updates a cell value in the specified sheet.

Parameters:
- `sheet_name`: Name of the sheet
- `cell_address`: Cell address in A1 notation (e.g., "A1", "B5")
- `value`: New value for the cell

## Requirements

- Python 3.8+
- openpyxl
- mcp