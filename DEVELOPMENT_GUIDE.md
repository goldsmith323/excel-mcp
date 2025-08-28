# Excel MCP Development Guide

This guide will help you continue developing the Excel MCP Integration system independently.

## 📁 Project Structure

```
excel-mcp/
├── src/excel_mcp/
│   ├── simple_server.py         # Main MCP server with engineering tools
│   ├── engineering_tools.py     # Advanced engineering analysis
│   ├── excel_tools.py          # Basic Excel operations
│   └── __init__.py             # Python package marker
├── excel_files/                # Sample Excel files for testing
├── excel_monitor_simple.py     # Desktop monitor application
├── start_simple_monitor.sh     # Launcher script
└── README.md                   # Project documentation
```

## 🚀 How It Works

1. **Excel Monitor** (`excel_monitor_simple.py`) detects when you open Excel files
2. **User chooses** to connect the file to MCP Desktop client
3. **MCP Server** (`simple_server.py`) is configured and started
4. **MCP Desktop client** connects to the MCP server
5. **System analyzes** the Excel file using engineering tools

## 🛠️ Development Areas

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
        text=f"✅ Result: {result}"
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

### Modifying Excel Monitor

**Location**: `excel_monitor_simple.py`

Key methods to modify:
- `get_open_excel_files()` - Change file detection logic
- `prompt_user_for_integration()` - Modify user prompts
- `update_claude_config()` - Change Claude Desktop configuration

### Supporting New File Types

Currently supports: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`

**To add new types:**
1. Update file extensions in `get_open_excel_files()`
2. Ensure `openpyxl` or add new library can read the format
3. Test with sample files

## 🔧 Testing Your Changes

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

### 3. Test in Claude Desktop
1. Connect a file using the monitor
2. Open Claude Desktop
3. Try commands like:
   - "Analyze this engineering calculator"
   - "Show me the input parameters"
   - "What formulas are used?"

## 📚 Key Libraries

### MCP (Model Context Protocol)
- **Purpose**: Communication between Claude Desktop and your server
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

## 🎯 Common Development Tasks

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

## 🐛 Debugging Tips

### MCP Server Issues
- Check error messages in Terminal where you started the monitor
- Verify `EXCEL_FILE_PATH` environment variable is set correctly
- Test MCP server directly with manual environment variable

### Excel Detection Issues
- Check if Excel process names match your system
- Verify file extensions are included in detection logic
- Test `get_open_excel_files()` method independently

### Claude Desktop Connection Issues
- Check `claude_desktop_config.json` is updated correctly
- Restart Claude Desktop after configuration changes
- Verify file paths are absolute and correct

## 📝 Code Style Guidelines

- **Use descriptive variable names** for engineering concepts
- **Add docstrings** to all public methods
- **Handle errors gracefully** with try/catch blocks
- **Format results consistently** for MCP client responses
- **Add type hints** where possible

## 🔄 Version Control Best Practices

- **Commit frequently** with descriptive messages
- **Test before committing** major changes
- **Keep backups** of working configurations
- **Document breaking changes** in commit messages

## 📞 Getting Help

If you need assistance:
1. **Check error messages** carefully - they often indicate the exact issue
2. **Test components individually** - MCP server, monitor, etc.
3. **Review MCP documentation** for protocol details
4. **Check OpenPyXL docs** for Excel file operations

## 🎓 Learning Resources

- **MCP Protocol**: https://modelcontextprotocol.io/docs/
- **Excel File Formats**: OpenPyXL documentation
- **Engineering Standards**: UFC, AISC, ACI documentation
- **Python Development**: Python.org tutorials

---

**Happy developing! 🚀**

This system provides a solid foundation for Excel-MCP integration with room for extensive customization and enhancement.