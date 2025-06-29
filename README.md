# Excel MCP Server

A Model Context Protocol (MCP) server for Excel file operations using Python.

## Features

- **get_document_info**: Get basic information about Excel documents including sheet names, dimensions, and metadata
- **update_cell**: Update cell values using A1 notation (e.g., "A1", "B5")

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the server:
```bash
python3 -m src.excel_mcp.server
```

## Usage

The server automatically uses `MasterBlaster v1.4.2.xlsx` in the current directory, or you can set the `EXCEL_FILE_PATH` environment variable:

```bash
EXCEL_FILE_PATH=/path/to/your/file.xlsx python3 -m src.excel_mcp.server
```

## Development

Test the Excel operations without the MCP server:
```bash
python3 test_excel.py
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