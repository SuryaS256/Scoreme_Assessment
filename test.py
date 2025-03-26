!pip install PyPDF2 pandas openpyxl

import re
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from google.colab import files

class PDFTableExtractor:
    def __init__(self):
        self.min_col_width = 3  # Minimum characters to consider as a column
        self.min_rows = 2  # Minimum rows to consider as a table
        self.header_threshold = 0.7  # Similarity threshold for header detection

    def extract_tables_from_pdf(self, pdf_path):
        """Extract tables from a PDF file"""
        tables = []
        
        # Read PDF file
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                text = page.extract_text()
                
                # Split text into lines
                lines = text.split('\n')
                
                # Find potential tables
                potential_tables = self._find_potential_tables(lines)
                
                # Process potential tables
                for table_lines in potential_tables:
                    table_data = self._parse_table(table_lines)
                    if table_data and len(table_data) >= self.min_rows:
                        tables.append({
                            'page': page_num + 1,
                            'data': table_data
                        })
        
        return tables

    def _find_potential_tables(self, lines):
        """Identify potential table sections in the text lines"""
        tables = []
        current_table = []
        in_table = False
        
        for line in lines:
            # Check if line has potential table structure
            if self._is_potential_table_row(line):
                if not in_table:
                    in_table = True
                current_table.append(line)
            else:
                if in_table and len(current_table) >= self.min_rows:
                    tables.append(current_table)
                in_table = False
                current_table = []
        
        # Add the last table if we're still in one
        if in_table and len(current_table) >= self.min_rows:
            tables.append(current_table)
            
        return tables

    def _is_potential_table_row(self, line):
        """Determine if a line looks like a table row"""
        # Check for common table patterns
        # 1. Multiple values separated by whitespace
        parts = re.split(r'\s{2,}', line.strip())
        if len(parts) >= 2:
            return True
            
        # 2. Pipe or other delimiter separated values
        if '|' in line and len(line.split('|')) >= 2:
            return True
            
        return False

    def _parse_table(self, lines):
        """Parse a table from lines of text"""
        # First pass: determine column boundaries
        column_boundaries = self._find_column_boundaries(lines)
        
        if not column_boundaries:
            return None
            
        # Second pass: extract data using column boundaries
        table_data = []
        for line in lines:
            row = []
            prev_boundary = 0
            for boundary in column_boundaries:
                cell = line[prev_boundary:boundary].strip()
                row.append(cell)
                prev_boundary = boundary
            # Add the last cell
            cell = line[prev_boundary:].strip()
            row.append(cell)
            
            table_data.append(row)
        
        return table_data

    def _find_column_boundaries(self, lines):
        """Find column boundaries by analyzing whitespace patterns"""
        if not lines:
            return None
            
        # Create a list to track whitespace runs
        whitespace_runs = []
        
        # Analyze each line for consistent whitespace patterns
        for line in lines:
            current_runs = []
            in_space = False
            space_start = 0
            
            for i, char in enumerate(line):
                if char.isspace():
                    if not in_space:
                        in_space = True
                        space_start = i
                else:
                    if in_space:
                        in_space = False
                        space_length = i - space_start
                        if space_length >= self.min_col_width:
                            current_runs.append((space_start, i))
            
            whitespace_runs.append(current_runs)
        
        # Find common boundaries across lines
        if not whitespace_runs:
            return None
            
        # Start with the first line's boundaries
        common_boundaries = [run[1] for run in whitespace_runs[0]]
        
        # Compare with other lines
        for runs in whitespace_runs[1:]:
            new_boundaries = []
            for boundary in common_boundaries:
                # Look for a similar boundary in this line
                for run in runs:
                    if abs(run[1] - boundary) <= 2:  # Allow small variations
                        new_boundaries.append(run[1])
                        break
            common_boundaries = new_boundaries
            
        return common_boundaries if common_boundaries else None

    def save_tables_to_excel(self, tables, output_path):
        """Save extracted tables to an Excel file"""
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet
        
        for i, table in enumerate(tables):
            # Create a new sheet for each table
            sheet_name = f"Page_{table['page']}_Table_{i+1}"
            ws = wb.create_sheet(title=sheet_name)
            
            # Convert table data to DataFrame
            df = pd.DataFrame(table['data'])
            
            # Write DataFrame to sheet
            for r in dataframe_to_rows(df, index=False, header=False):
                ws.append(r)
        
        wb.save(output_path)

def process_specific_pdf():
    """Process the specific PDF file at /content/pdf_extract/test3 (1).pdf"""
    # Define file paths
    pdf_path = "/content/pdf_extract/test3 (1).pdf"
    output_excel = "/content/test3_tables.xlsx"
    
    # Verify file exists
    if not os.path.exists(pdf_path):
        print(f"Error: File not found at {pdf_path}")
        print("Please ensure:")
        print("1. The file exists in the specified location")
        print("2. You've mounted Google Drive if the file is there")
        print("3. The path is correct (including the space before (1))")
        return
    
    # Process the PDF
    extractor = PDFTableExtractor()
    tables = extractor.extract_tables_from_pdf(pdf_path)
    
    if tables:
        extractor.save_tables_to_excel(tables, output_excel)
        print(f"Successfully extracted {len(tables)} tables to {output_excel}")
        
        # Download the result
        files.download(output_excel)
    else:
        print("No tables found in the PDF")

# Run the processor
process_specific_pdf()
