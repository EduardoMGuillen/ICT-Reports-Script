import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import glob
from datetime import datetime

def parse_csv_report(file_path):
    """
    Parse the CSV report file and return a cleaned DataFrame
    """
    try:
        # Read the CSV file with proper handling
        df = pd.read_csv(file_path, skipinitialspace=True)
        
        # Clean column names by stripping whitespace
        df.columns = df.columns.str.strip()
        
        # Ensure we have the expected columns
        expected_columns = ['Result', 'Board', 'Type', 'Part', 'ActVal', 'StdVal', 'HL', 'LL', 'Mode', 'Range', 'TestVal']
        
        # Check if all expected columns exist
        missing_columns = [col for col in expected_columns if col not in df.columns]
        if missing_columns:
            print(f"Warning: Missing columns in {file_path}: {missing_columns}")
        
        return df
        
    except Exception as e:
        print(f"Error reading CSV file {file_path}: {e}")
        return None

def extract_product_info(filename):
    """
    Extract product information from filename
    """
    # Extract date from filename (assuming format like 20250627_EBP4742PASS23675AC2516801315F3.TXT)
    try:
        date_part = filename[:8]
        formatted_date = f"{date_part[6:8]}-{get_month_name(date_part[4:6])}-{date_part[2:4]}"
        
        # Extract program info (this might need adjustment based on your naming convention)
        program = "AUT1077-HC00"  # Default, you can modify this logic
        smt_line = "Line 9"       # Default, you can modify this logic
        
        return {
            'Program': program,
            'SMT Run': formatted_date,
            'SMT Line': smt_line
        }
    except:
        return {
            'Program': 'AUT1083-HC00',
            'SMT Run': datetime.now().strftime('%d-%b-%y'),
            'SMT Line': 'Line 9'
        }

def get_month_name(month_num):
    """Convert month number to abbreviated name"""
    months = {
        '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr',
        '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug',
        '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
    }
    return months.get(month_num, 'Jun')

def create_formatted_excel(df, product_info, output_path):
    """
    Create a formatted Excel file matching the template exactly
    """
    # Create a new workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Test Report"
    
    # Remove gridlines
    ws.sheet_view.showGridLines = False
    
    # Define styles
    header_font = Font(name='Arial', size=11, bold=True)
    data_font = Font(name='Arial', size=10)
    center_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Header colors
    header_fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
    pass_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    
    # Title
    ws['A1'] = 'Product Information'
    ws['A1'].font = Font(name='Arial', size=14, bold=True)
    ws.merge_cells('A1:B1')
    
    # Product Information Section
    info_data = [
        ['Program', product_info['Program']],
        ['SMT Run', product_info['SMT Run']],
        ['SMT Line', product_info['SMT Line']]
    ]
    
    for i, (key, value) in enumerate(info_data, start=2):
        ws[f'A{i}'] = key
        ws[f'B{i}'] = value
        ws[f'A{i}'].font = header_font
        ws[f'A{i}'].fill = header_fill
        ws[f'A{i}'].border = thin_border
        ws[f'A{i}'].alignment = center_alignment
        ws[f'B{i}'].font = data_font
        ws[f'B{i}'].border = thin_border
        ws[f'B{i}'].alignment = center_alignment
    
    # Legend Section - exact layout matching the second image (no gridlines)
    # Row 7: Result | Result of the measured | HL | High Rate Tolerance
    ws['A7'] = 'Result'
    ws['A7'].font = header_font
    ws['A7'].fill = header_fill
    ws['A7'].border = thin_border
    ws['A7'].alignment = center_alignment
    
    ws['B7'] = 'Result of the measured'
    ws['B7'].font = data_font
    ws['B7'].border = thin_border
    ws['B7'].alignment = center_alignment
    
    ws['C7'] = 'HL'
    ws['C7'].font = header_font
    ws['C7'].fill = header_fill
    ws['C7'].border = thin_border
    ws['C7'].alignment = center_alignment
    
    ws['D7'] = 'High Rate Tolerance'
    ws['D7'].font = data_font
    ws['D7'].border = thin_border
    ws['D7'].alignment = center_alignment
    
    # Row 8: Board | Module of the board | LL | Low Rate Tolerance
    ws['A8'] = 'Board'
    ws['A8'].font = header_font
    ws['A8'].fill = header_fill
    ws['A8'].border = thin_border
    ws['A8'].alignment = center_alignment
    
    ws['B8'] = 'Module of the board'
    ws['B8'].font = data_font
    ws['B8'].border = thin_border
    ws['B8'].alignment = center_alignment
    
    ws['C8'] = 'LL'
    ws['C8'].font = header_font
    ws['C8'].fill = header_fill
    ws['C8'].border = thin_border
    ws['C8'].alignment = center_alignment
    
    ws['D8'] = 'Low Rate Tolerance'
    ws['D8'].font = data_font
    ws['D8'].border = thin_border
    ws['D8'].alignment = center_alignment
    
    # Row 9: Type | Type of component | Mode | Type of measure: CC - Constant current, AC - Alternate current, LV - Low voltage
    ws['A9'] = 'Type'
    ws['A9'].font = header_font
    ws['A9'].fill = header_fill
    ws['A9'].border = thin_border
    ws['A9'].alignment = center_alignment
    
    ws['B9'] = 'Type of component'
    ws['B9'].font = data_font
    ws['B9'].border = thin_border
    ws['B9'].alignment = center_alignment
    
    ws['C9'] = 'Mode'
    ws['C9'].font = header_font
    ws['C9'].fill = header_fill
    ws['C9'].border = thin_border
    ws['C9'].alignment = center_alignment
    
    ws['D9'] = 'Type of measure: CC - Constant current,\nAC - Alternate current,\nLV - Low voltage'
    ws['D9'].font = data_font
    ws['D9'].border = thin_border
    ws['D9'].alignment = center_alignment
    
    # Row 10: Part | Name of the component | Range | Range of measured
    ws['A10'] = 'Part'
    ws['A10'].font = header_font
    ws['A10'].fill = header_fill
    ws['A10'].border = thin_border
    ws['A10'].alignment = center_alignment
    
    ws['B10'] = 'Name of the component'
    ws['B10'].font = data_font
    ws['B10'].border = thin_border
    ws['B10'].alignment = center_alignment
    
    ws['C10'] = 'Range'
    ws['C10'].font = header_font
    ws['C10'].fill = header_fill
    ws['C10'].border = thin_border
    ws['C10'].alignment = center_alignment
    
    ws['D10'] = 'Range of measured'
    ws['D10'].font = data_font
    ws['D10'].border = thin_border
    ws['D10'].alignment = center_alignment
    
    # Row 10: ActVal | Value paralel components | TestVal | Real measured value
    ws['A11'] = 'ActVal'
    ws['A11'].font = header_font
    ws['A11'].fill = header_fill
    ws['A11'].border = thin_border
    ws['A11'].alignment = center_alignment
    
    ws['B11'] = 'Value paralel\ncomponents'
    ws['B11'].font = data_font
    ws['B11'].border = thin_border
    ws['B11'].alignment = center_alignment
    
    ws['C11'] = 'TestVal'
    ws['C11'].font = header_font
    ws['C11'].fill = header_fill
    ws['C11'].border = thin_border
    ws['C11'].alignment = center_alignment
    
    ws['D11'] = 'Real measured value'
    ws['D11'].font = data_font
    ws['D11'].border = thin_border
    ws['D11'].alignment = center_alignment
    
    # Row 11: StdVal | BOM value components | [Span across C12:D12 - empty]
    ws['A12'] = 'StdVal'
    ws['A12'].font = header_font
    ws['A12'].fill = header_fill
    ws['A12'].border = thin_border
    ws['A12'].alignment = center_alignment
    
    ws['B12'] = 'BOM value components'
    ws['B12'].font = data_font
    ws['B12'].border = thin_border
    ws['B12'].alignment = center_alignment
    
    ws['C12'] = ''
    ws['C12'].border = thin_border
    ws.merge_cells('C12:D12')
    
    # Data section starts after legend
    data_start_row = 14  # Starting at row 14 for data headers
    
    # Headers for data table
    headers = ['Result', 'Board', 'Type', 'Part', 'ActVal', 'StdVal', 'HL', 'LL', 'Mode', 'Range', 'TestVal']
    
    for j, header in enumerate(headers):
        col_letter = chr(65 + j)
        cell = ws[f'{col_letter}{data_start_row}']
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = center_alignment
    
    # Add data rows
    for i, row in df.iterrows():
        row_num = data_start_row + 1 + i
        result = str(row['Result']).strip().upper()
        
        for j, header in enumerate(headers):
            col_letter = chr(65 + j)
            cell = ws[f'{col_letter}{row_num}']
            cell.value = row[header] if header in row else ''
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = center_alignment
            
            # Color coding - ONLY the Result column (column A) with PASS values should be green
            if j == 0 and result == 'PASS':  # j == 0 means column A (Result column)
                cell.fill = pass_fill
    
    # Adjust column widths to match the image
    column_widths = {
        'A': 8,   # Result
        'B': 8,   # Board
        'C': 12,  # Type
        'D': 15,  # Part (wider for component names)
        'E': 10,  # ActVal
        'F': 10,  # StdVal
        'G': 8,   # HL
        'H': 8,   # LL
        'I': 8,   # Mode
        'J': 8,   # Range
        'K': 12   # TestVal
    }
    
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width
    
    # Save the workbook
    wb.save(output_path)
    print(f"Excel report saved to: {output_path}")

def process_all_files():
    """
    Process all TXT files in the input directory and create Excel reports
    """
    # Define input and output directories
    input_dir = r"C:\Users\eamaldonado\Desktop\ICT Report Script\TXT Files"
    output_dir = r"C:\Users\eamaldonado\Desktop\ICT Report Script\Excel Reports"
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Find all TXT files in the input directory
    txt_files = glob.glob(os.path.join(input_dir, "*.TXT")) + glob.glob(os.path.join(input_dir, "*.txt"))
    
    if not txt_files:
        print(f"No TXT files found in {input_dir}")
        return
    
    print(f"Found {len(txt_files)} TXT files to process:")
    for file in txt_files:
        print(f"  - {os.path.basename(file)}")
    
    processed_count = 0
    failed_count = 0
    
    for txt_file in txt_files:
        try:
            print(f"\nProcessing: {os.path.basename(txt_file)}")
            
            # Parse the CSV data
            df = parse_csv_report(txt_file)
            
            if df is None:
                print(f"Failed to read data from {os.path.basename(txt_file)}")
                failed_count += 1
                continue
            
            print(f"Successfully read {len(df)} records.")
            
            # Extract product information
            product_info = extract_product_info(os.path.basename(txt_file))
            
            # Create output filename based on input filename
            base_name = os.path.splitext(os.path.basename(txt_file))[0]
            output_file = os.path.join(output_dir, f"Test_Report_{base_name}.xlsx")
            
            # Create the formatted Excel file
            create_formatted_excel(df, product_info, output_file)
            
            # Display summary statistics
            result_counts = df['Result'].value_counts()
            print(f"Summary for {os.path.basename(txt_file)}:")
            for result, count in result_counts.items():
                print(f"  {result}: {count}")
            
            processed_count += 1
            
        except Exception as e:
            print(f"Error processing {os.path.basename(txt_file)}: {e}")
            failed_count += 1
    
    print(f"\n{'='*50}")
    print(f"BATCH PROCESSING COMPLETE")
    print(f"{'='*50}")
    print(f"Successfully processed: {processed_count} files")
    print(f"Failed to process: {failed_count} files")
    print(f"Output directory: {output_dir}")

def main():
    """
    Main function to process all TXT files and create Excel reports
    """
    print("ICT Report Batch Converter")
    print("=" * 50)
    
    try:
        process_all_files()
    except Exception as e:
        print(f"An error occurred during batch processing: {e}")
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()