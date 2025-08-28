# Excel MCP Server

A Model Context Protocol (MCP) server for Excel file operations and engineering calculator analysis using Python.

**Author:** Hossein Zargar

## Features

- **get_document_info**: Get basic information about Excel documents including sheet names, dimensions, and metadata
- **update_cell**: Update cell values using A1 notation (e.g., "A1", "B5")
- **Advanced Engineering Analysis**: Analyze engineering calculators with parameter detection and formula analysis
- **Cross-platform**: Works on Windows, macOS, and Linux
- **Sample Files**: Includes sample Excel files for testing

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

## Project Structure

```
excel-mcp/
├── src/excel_mcp/
│   ├── simple_server.py         # Main MCP server with engineering tools
│   ├── engineering_tools.py     # Advanced engineering analysis
│   ├── excel_tools.py          # Basic Excel operations
│   └── __init__.py             # Python package marker
├── excel_files/                # Sample Excel files for testing
├── excel_monitor_simple.py     # Desktop monitor application
├── start_simple_monitor.sh     # Launcher script (Linux/macOS)
├── start_simple_monitor.bat    # Launcher script (Windows)
└── README.md                   # Project documentation
```

## How It Works

1. **Excel Monitor** (`excel_monitor_simple.py`) detects when you open Excel files
2. **User chooses** to connect the file to MCP Desktop client
3. **MCP Server** (`simple_server.py`) is configured and started
4. **MCP Desktop client** connects to the MCP server
5. **System analyzes** the Excel file using engineering tools

## MCP Tools

### Basic Excel Operations

#### get_document_info
Returns information about the Excel file including:
- File path and size
- Number of sheets
- Sheet names and dimensions

#### update_cell
Updates a cell value in the specified sheet.

Parameters:
- `sheet_name`: Name of the sheet
- `cell_address`: Cell address in A1 notation (e.g., "A1", "B5")
- `value`: New value for the cell

### Engineering Analysis Tools

#### analyze_engineering_calculator
Performs comprehensive analysis of engineering calculators including:
- Calculator type identification
- Input/output parameter detection
- Formula complexity analysis
- Units and validation rules

#### find_input_parameters
Locates and analyzes input parameters in engineering calculators.

#### get_calculator_summary
Provides a high-level summary of calculator capabilities and engineering domain.

## Development Guide

### Adding New MCP Tools

**Location**: `src/excel_mcp/simple_server.py`

1. **Add tool definition** in `handle_list_tools()`:
```python
types.Tool(
    name="your_new_tool",
    description="What your tool does",
    inputSchema={
        "type": "object",
        "properties": {
            "param_name": {
                "type": "string",
                "description": "Parameter description"
            }
        },
        "required": ["param_name"]
    }
)
```

2. **Add tool handler** in `handle_call_tool()`:
```python
elif name == "your_new_tool":
    param_value = arguments.get("param_name")
    
    # Your tool logic here
    result = do_something(param_value)
    
    return [types.TextContent(
        type="text",
        text=f"Result: {result}"
    )]
```

### Extending Engineering Analysis

**Location**: `src/excel_mcp/engineering_tools.py`

The `EngineeringExcelAnalyzer` class contains methods for:
- `_identify_calculator_type()` - Determine calculator purpose
- `_find_input_parameters()` - Locate input fields
- `_analyze_formulas()` - Examine Excel formulas
- `_analyze_units()` - Check units and conversions

**To add new analysis:**
1. Add method to `EngineeringExcelAnalyzer` class
2. Call it from `analyze_calculator_structure()`
3. Add corresponding MCP tool in `simple_server.py`

### Supporting New File Types

Currently supports: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`

**To add new types:**
1. Update file extensions in `get_open_excel_files()`
2. Ensure `openpyxl` or add new library can read the format
3. Test with sample files

## Testing

### 1. Test MCP Server Directly

**On Linux/macOS:**
```bash
cd excel-mcp
EXCEL_FILE_PATH="/path/to/test.xlsx" python3 src/excel_mcp/simple_server.py
```

**On Windows:**
```cmd
cd excel-mcp
set EXCEL_FILE_PATH=excel_files\beam_analysis.xlsx
python src/excel_mcp/simple_server.py
```

### 2. Test with Excel Monitor

**On Linux/macOS:**
```bash
cd excel-mcp
./start_simple_monitor.sh
```

**On Windows:**
```cmd
cd excel-mcp
start_simple_monitor.bat
```

Then open an Excel file and choose to connect when prompted.

### 3. Test with MCP Client
1. Connect a file using the monitor
2. Open your MCP client
3. Try commands like:
   - "Analyze this engineering calculator"
   - "Show me the input parameters"
   - "What formulas are used?"

## Key Libraries

### MCP (Model Context Protocol)
- **Purpose**: Communication between MCP clients and your server
- **Docs**: https://modelcontextprotocol.io/
- **Key concepts**: Tools, Resources, Prompts

### OpenPyXL
- **Purpose**: Read/write Excel files
- **Docs**: https://openpyxl.readthedocs.io/
- **Key features**: Cell access, formulas, formatting

### psutil
- **Purpose**: System and process monitoring
- **Docs**: https://psutil.readthedocs.io/
- **Use case**: Detect running Excel processes

## Common Development Tasks

### Add Support for New Engineering Domain

1. **Update domain detection** in `_identify_calculator_type()`:
```python
elif any(word in sheet_text for word in ['your', 'domain', 'keywords']):
    calculator_info['engineering_domain'] = 'Your Engineering Domain'
```

2. **Add domain-specific patterns** in `engineering_keywords`:
```python
'your_domain': ['keyword1', 'keyword2', 'keyword3']
```

### Improve Parameter Detection

1. **Add new parameter patterns** in `engineering_keywords`
2. **Enhance `_analyze_parameter_cell()`** to look in more locations
3. **Update unit patterns** in `unit_patterns`

### Add Formula Validation

1. **Create new method** in `EngineeringExcelAnalyzer`:
```python
def _validate_formulas(self) -> List[Dict[str, Any]]:
    # Your validation logic
    pass
```

2. **Add MCP tool** for formula validation
3. **Test with engineering calculators**

## Debugging Tips

### MCP Server Issues
- Check error messages in Terminal where you started the monitor
- Verify `EXCEL_FILE_PATH` environment variable is set correctly
- Test MCP server directly with manual environment variable

### Excel Detection Issues
- Check if Excel process names match your system
- Verify file extensions are included in detection logic
- Test `get_open_excel_files()` method independently

### MCP Client Connection Issues
- Check configuration is updated correctly
- Restart MCP client after configuration changes
- Verify file paths are absolute and correct

## Code Style Guidelines

- **Use descriptive variable names** for engineering concepts
- **Add docstrings** to all public methods
- **Handle errors gracefully** with try/catch blocks
- **Format results consistently** for MCP client responses
- **Add type hints** where possible

## Requirements

- Python 3.8+
- openpyxl>=3.1.0
- psutil>=5.9.0
- mcp>=1.0.0
- pandas>=2.0.0
- numpy>=1.24.0

## License

MIT License - see LICENSE file for details.

## Learning Resources

- **MCP Protocol**: https://modelcontextprotocol.io/docs/
- **Excel File Formats**: OpenPyXL documentation
- **Engineering Standards**: UFC, AISC, ACI documentation
- **Python Development**: Python.org tutorials

---

This system provides a solid foundation for Excel-MCP integration with room for extensive customization and enhancement.