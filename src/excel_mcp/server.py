"""Excel MCP Server implementation."""

import asyncio
import os
from typing import Any, Sequence
from mcp.server import Server
from mcp.server.models import InitializationOptions
from mcp.server.stdio import stdio_server
from mcp.types import (
    Resource,
    Tool,
    TextContent,
    ImageContent,
    EmbeddedResource,
    LoggingLevel
)
import mcp.server.stdio
import mcp.types as types

from .excel_tools import ExcelHandler


# Initialize the MCP server
server = Server("excel-mcp")

# Global Excel handler - will be initialized with file path
excel_handler: ExcelHandler = None


@server.list_tools()
async def handle_list_tools() -> list[Tool]:
    """List available tools."""
    return [
        Tool(
            name="get_document_info",
            description="Get basic information about the Excel document including sheet names, dimensions, and metadata",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        Tool(
            name="update_cell",
            description="Update a cell value in the Excel document",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the sheet to update"
                    },
                    "cell_address": {
                        "type": "string",
                        "description": "Cell address in A1 notation (e.g., 'A1', 'B5', 'C10')"
                    },
                    "value": {
                        "description": "Value to set in the cell (can be string, number, or formula)"
                    }
                },
                "required": ["sheet_name", "cell_address", "value"]
            }
        )
    ]


@server.call_tool()
async def handle_call_tool(name: str, arguments: dict) -> list[types.TextContent]:
    """Handle tool calls."""
    if excel_handler is None:
        return [types.TextContent(
            type="text",
            text="Error: Excel file not initialized. Please set EXCEL_FILE_PATH environment variable."
        )]
    
    try:
        if name == "get_document_info":
            result = excel_handler.get_document_info()
            return [types.TextContent(
                type="text",
                text=f"Excel Document Information:\n\n"
                     f"File: {result['file_path']}\n"
                     f"Size: {result['file_size']} bytes\n"
                     f"Number of sheets: {result['sheet_count']}\n\n"
                     f"Sheets:\n" + 
                     "\n".join([f"  - {sheet['name']}: {sheet['dimensions']} (max row: {sheet['max_row']}, max col: {sheet['max_column']})" 
                               for sheet in result['sheets']])
            )]
        
        elif name == "update_cell":
            sheet_name = arguments.get("sheet_name")
            cell_address = arguments.get("cell_address")
            value = arguments.get("value")
            
            result = excel_handler.update_cell(sheet_name, cell_address, value)
            return [types.TextContent(
                type="text",
                text=f"Cell Update Result:\n\n"
                     f"Sheet: {result['sheet']}\n"
                     f"Cell: {result['cell']}\n"
                     f"Old value: {result['old_value']}\n"
                     f"New value: {result['new_value']}\n"
                     f"Status: {'Success' if result['success'] else 'Failed'}"
            )]
        
        else:
            return [types.TextContent(
                type="text",
                text=f"Unknown tool: {name}"
            )]
    
    except Exception as e:
        return [types.TextContent(
            type="text",
            text=f"Error executing {name}: {str(e)}"
        )]


async def main():
    """Main entry point for the Excel MCP server."""
    global excel_handler
    
    # Get Excel file path from environment or use default
    excel_file_path = os.getenv("EXCEL_FILE_PATH", "MasterBlaster v1.4.2.xlsx")
    
    # Make path absolute if it's relative
    if not os.path.isabs(excel_file_path):
        excel_file_path = os.path.abspath(excel_file_path)
    
    try:
        excel_handler = ExcelHandler(excel_file_path)
        print(f"Excel MCP Server initialized with file: {excel_file_path}", file=os.sys.stderr)
    except FileNotFoundError as e:
        print(f"Error: {e}", file=os.sys.stderr)
        print(f"Please ensure the Excel file exists or set EXCEL_FILE_PATH environment variable.", file=os.sys.stderr)
        exit(1)
    
    # Run the server
    async with mcp.server.stdio.stdio_server(server) as (read_stream, write_stream):
        await server.run(
            read_stream, write_stream, InitializationOptions(
                server_name="excel-mcp",
                server_version="0.1.0",
                capabilities=server.get_capabilities(
                    notification_options=None,
                    experimental_capabilities=None,
                )
            )
        )


if __name__ == "__main__":
    asyncio.run(main())