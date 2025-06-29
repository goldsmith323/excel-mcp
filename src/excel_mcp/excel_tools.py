"""Excel operations module."""

import os
from typing import Dict, Any, List
from openpyxl import load_workbook
from openpyxl.workbook import Workbook


class ExcelHandler:
    """Handles Excel file operations."""
    
    def __init__(self, file_path: str):
        """Initialize with Excel file path."""
        self.file_path = file_path
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found: {file_path}")
    
    def get_document_info(self) -> Dict[str, Any]:
        """Get basic information about the Excel document."""
        workbook = load_workbook(self.file_path)
        
        sheet_info = []
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            max_row = sheet.max_row
            max_col = sheet.max_column
            sheet_info.append({
                "name": sheet_name,
                "max_row": max_row,
                "max_column": max_col,
                "dimensions": f"{max_row}x{max_col}"
            })
        
        return {
            "file_path": self.file_path,
            "file_size": os.path.getsize(self.file_path),
            "sheet_count": len(workbook.sheetnames),
            "sheets": sheet_info
        }
    
    def update_cell(self, sheet_name: str, cell_address: str, value: Any) -> Dict[str, Any]:
        """Update a cell value in the specified sheet."""
        workbook = load_workbook(self.file_path)
        
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}")
        
        sheet = workbook[sheet_name]
        
        # Update the cell
        old_value = sheet[cell_address].value
        sheet[cell_address] = value
        
        # Save the workbook
        workbook.save(self.file_path)
        
        return {
            "sheet": sheet_name,
            "cell": cell_address,
            "old_value": old_value,
            "new_value": value,
            "success": True
        }