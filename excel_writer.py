# excel_writer.py
"""Excel export functionality."""

import logging
from pathlib import Path
from typing import List
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment
from models import Document

class ExcelWriter:
    """Handles writing document data to Excel files."""
    
    def __init__(self, output_path: Path):
        self.output_path = output_path
        self.workbook = Workbook()
        self.logger = logging.getLogger(__name__)
        
        # Remove default sheet
        if 'Sheet' in self.workbook.sheetnames:
            self.workbook.remove(self.workbook['Sheet'])
    
    def write_summary(self, documents: List[Document]):
        """Write summary table to Excel."""
        # Create summary data
        summary_data = []
        for doc in documents:
            summary_data.append({
                'ID': doc.id,
                'Parent Folder': doc.parent_folder,
                'Document Name': doc.name,
                'Document Title': doc.title,
                'Word Count': doc.word_count,
                'Image Count': doc.image_count,
                'Unique Image Count': doc.unique_image_count
            })
        
        # Create DataFrame
        df = pd.DataFrame(summary_data)
        
        # Create worksheet
        ws = self.workbook.create_sheet("Summary")
        
        # Write data
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Format header
        self._format_header(ws)
        
        # Auto-adjust column widths
        self._adjust_column_widths(ws)
        
        self.logger.info(f"Written summary for {len(documents)} documents")
    
    def write_sections(self, documents: List[Document]):
        """Write section details for each document."""
        for doc in documents:
            if not doc.sections:
                continue
            
            # Create sheet name (limit to 31 chars for Excel)
            sheet_name = f"Doc_{doc.id}_{doc.filename[:20]}"
            sheet_name = self._sanitize_sheet_name(sheet_name)
            
            # Create worksheet
            ws = self.workbook.create_sheet(sheet_name)
            
            # Add document info
            ws.append(['Document ID:', doc.id])
            ws.append(['Document Name:', doc.name])
            ws.append(['Document Title:', doc.title])
            ws.append([])  # Empty row
            
            # Add headers
            ws.append(['Section Heading', 'Font Name', 'Font Size', 'Text Preview'])
            
            # Format section header row
            header_row = ws.max_row
            for cell in ws[header_row]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
            # Add section data
            for section in doc.sections:
                # Limit text preview to 200 characters
                text_preview = section.text[:200] + "..." if len(section.text) > 200 else section.text
                text_preview = text_preview.replace('\n', ' ')  # Remove newlines
                
                ws.append([
                    section.heading,
                    section.font_name or "Default",
                    section.font_size or "Default",
                    text_preview
                ])
            
            # Format document info cells
            for row in range(1, 4):
                ws.cell(row=row, column=1).font = Font(bold=True)
            
            # Auto-adjust column widths
            self._adjust_column_widths(ws)
            
            self.logger.info(f"Written {len(doc.sections)} sections for document {doc.id}")
    
    def save(self):
        """Save the workbook to file."""
        self.workbook.save(self.output_path)
        self.logger.info(f"Excel file saved to: {self.output_path}")
    
    def _format_header(self, worksheet):
        """Format header row with bold font and background color."""
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")
    
    def _adjust_column_widths(self, worksheet):
        """Auto-adjust column widths based on content."""
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    def _sanitize_sheet_name(self, name: str) -> str:
        """Sanitize sheet name for Excel compatibility."""
        # Remove invalid characters
        invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
        for char in invalid_chars:
            name = name.replace(char, '_')
        
        # Limit to 31 characters
        return name[:31]
