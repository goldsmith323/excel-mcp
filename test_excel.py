#!/usr/bin/env python3
"""Simple test script for Excel operations."""

from src.excel_mcp.excel_tools import ExcelHandler

def main():
    try:
        # Test with the Excel file
        handler = ExcelHandler("MasterBlaster v1.4.2.xlsx")
        
        print("=== Testing get_document_info ===")
        info = handler.get_document_info()
        print(f"File: {info['file_path']}")
        print(f"Size: {info['file_size']} bytes")
        print(f"Sheets: {info['sheet_count']}")
        for sheet in info['sheets']:
            print(f"  - {sheet['name']}: {sheet['dimensions']}")
        
        print("\n=== Testing update_cell ===")
        # Get the first sheet name
        first_sheet = info['sheets'][0]['name']
        
        # Update a cell (A1)
        result = handler.update_cell(first_sheet, "A1", "Hello from MCP!")
        print(f"Updated {result['sheet']}:{result['cell']}")
        print(f"Old value: {result['old_value']}")
        print(f"New value: {result['new_value']}")
        
        print("\nTest completed successfully!")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()