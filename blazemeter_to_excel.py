#!/usr/bin/env python3
"""
BlazeMeter CSV to Excel Converter

This script converts BlazeMeter test result CSV files to Performance Test Result Template format.
It extracts summary data from the "ALL" row and creates a formatted report.

Usage:
    python blazemeter_to_excel.py <input_csv_file> <test_type> [output_xlsx_file]
    
    test_type: API or UI
    - API: P95% > 500ms = Fail (red), P95% <= 500ms = Pass (blue)
    - UI: P95% > 2000ms = Fail (red), P95% <= 2000ms = Pass (blue)

Example:
    python blazemeter_to_excel.py data.csv API
    # Creates: data-API-converted-YYYYMMDD-HHMMSS.xlsx
    
    python blazemeter_to_excel.py data.csv UI output.xlsx
    # Creates: output.xlsx
"""

import sys
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.numbers import FORMAT_NUMBER_00
from openpyxl.utils import get_column_letter
from datetime import datetime
import shutil
import glob


def convert_blazemeter_to_excel(input_csv, test_type='API', output_xlsx=None):
    """
    Convert BlazeMeter CSV to Performance Test Result Template format.
    
    Args:
        input_csv (str): Path to input CSV file
        test_type (str): Test type - 'API' or 'UI'
        output_xlsx (str, optional): Path to output Excel file. 
                                     If None, uses input filename with timestamp
    
    Returns:
        str: Path to the created Excel file
    """
    # Validate input file exists
    if not os.path.exists(input_csv):
        raise FileNotFoundError(f"Input file not found: {input_csv}")
    
    # Determine output filename with test type and timestamp
    if output_xlsx is None:
        base_name = os.path.splitext(input_csv)[0]
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        output_xlsx = f"{base_name}-{test_type.upper()}-converted-{timestamp}.xlsx"
    
    print(f"\nReading CSV file: {input_csv}")
    
    # Read the CSV file
    df = pd.read_csv(input_csv)
    
    # Extract the "ALL" row data for summary
    all_row = df[df['Element Label'] == 'ALL']
    
    if all_row.empty:
        print(f"  ⚠ No 'ALL' row found. Calculating aggregated data from all transactions...")
        
        # Calculate ALL row data from all transactions
        total_hits = int(df['# Samples'].sum())
        avg_bandwidth = df['Avg. Bandwidth (KBytes/s)'].sum()
        avg_hits_per_sec = df['Avg. Hits/s'].mean()
        
        # Calculate weighted average response time based on # Samples
        if total_hits > 0:
            avg_response_time = (df['Avg. Response Time (ms)'] * df['# Samples']).sum() / total_hits
        else:
            avg_response_time = 0
        
        # Calculate weighted error percentage based on # Samples
        if total_hits > 0:
            errors = (df['Error Percentage'] * df['# Samples']).sum() / total_hits
        else:
            errors = 0
        
        # Calculate percentiles for response times using weighted average based on # Samples
        # Use the actual percentile values from each transaction
        total_samples = df['# Samples'].sum()
        
        if total_samples > 0:
            # Calculate weighted percentiles
            if '90% line (ms)' in df.columns:
                response_90 = int((df['90% line (ms)'] * df['# Samples']).sum() / total_samples)
            else:
                response_90 = 0
            
            if '95% line (ms)' in df.columns:
                response_95 = int((df['95% line (ms)'] * df['# Samples']).sum() / total_samples)
            else:
                response_95 = 0
            
            if '99% line (ms)' in df.columns:
                response_99 = int((df['99% line (ms)'] * df['# Samples']).sum() / total_samples)
            else:
                response_99 = 0
        else:
            response_90 = 0
            response_95 = 0
            response_99 = 0
        
        # Get min and max response times
        min_response = df['Min Response Time (ms)'].min() if 'Min Response Time (ms)' in df.columns else 0
        max_response = df['Max Response Time (ms)'].max() if 'Max Response Time (ms)' in df.columns else 0
        
        print(f"  ✓ Calculated aggregated summary data from {len(df)} transactions")
    else:
        # Extract summary values from existing ALL row
        avg_bandwidth = all_row['Avg. Bandwidth (KBytes/s)'].values[0]
        total_hits = int(all_row['# Samples'].values[0])
        avg_hits_per_sec = all_row['Avg. Hits/s'].values[0]
        response_95 = int(all_row['95% line (ms)'].values[0])
        errors = all_row['Error Percentage'].values[0]
        avg_response_time = all_row['Avg. Response Time (ms)'].values[0]
        
        print(f"  ✓ Extracted summary data from 'ALL' row")
    
    # Remove the ALL row from transaction data
    df_transactions = df[df['Element Label'] != 'ALL'].copy()
    
    # Sort by Element Label in ascending order
    df_transactions = df_transactions.sort_values(by='Element Label', ascending=True)
    print(f"  ✓ Sorted transaction data by 'Element Label' (ascending)")
    
    # Define the desired column order for transactions
    column_order = [
        'Element Label',
        '# Samples',
        'Avg. Response Time (ms)',
        '95% line (ms)',
        'Error Percentage',
        'Avg. Hits/s',
        '90% line (ms)',
        '99% line (ms)',
        'Min Response Time (ms)',
        'Max Response Time (ms)',
        'Avg. Bandwidth (KBytes/s)',
        'Concurrency'
    ]
    
    # Reorder columns
    existing_columns = [col for col in column_order if col in df_transactions.columns]
    df_transactions = df_transactions[existing_columns]
    
    # Check if template exists
    template_path = '/Users/jameskim/Documents/Scripts/Python/Performance-Test-result-Template.xlsx'
    
    if os.path.exists(template_path):
        print(f"  ✓ Using template: {template_path}")
        # Copy template to output location
        shutil.copy(template_path, output_xlsx)
        workbook = load_workbook(output_xlsx)
        worksheet = workbook['Performance Test Report']
        
        # Unmerge all cells in the worksheet to allow writing
        merged_cells_to_unmerge = list(worksheet.merged_cells)
        
        for merged_cell in merged_cells_to_unmerge:
            worksheet.unmerge_cells(str(merged_cell))
        
        print(f"  ✓ Unmerged all cells for data update")
        
        # Update Row 1: Test name (should be the input filename without extension)
        test_name = os.path.splitext(os.path.basename(input_csv))[0]
        worksheet['B1'] = test_name
        
        # Re-merge Row 1 cells B1:F1 after writing
        worksheet.merge_cells('B1:F1')
        
        # Update Row 2: Date (should be empty)
        worksheet['B2'] = ''
        
        # Re-merge Row 2 cells B2:F2 after clearing
        worksheet.merge_cells('B2:F2')
        
        print(f"  ✓ Updated Test name (Row 1) with filename: {test_name}")
        print(f"  ✓ Cleared Date field (Row 2)")
        
        # Update summary data in the template (columns B-F, rows 3-7)
        # Row 3: Average Throughput
        worksheet['B3'] = avg_bandwidth
        worksheet.merge_cells('B3:F3')
        
        # Row 4: Total Hits
        worksheet['B4'] = total_hits
        worksheet.merge_cells('B4:F4')
        
        # Row 5: Average Hits per Second
        worksheet['B5'] = avg_hits_per_sec
        worksheet.merge_cells('B5:F5')
        
        # Row 6: 95% Response Time (P95% only)
        worksheet['B6'] = response_95
        worksheet.merge_cells('B6:F6')
        
        # Row 7: Errors (display as percentage with 2 decimal places)
        worksheet['B7'] = f"{errors:.2f}%"  # Format as text with % sign (e.g., "0.46%")
        worksheet.merge_cells('B7:F7')
        
        print(f"  ✓ Updated summary data in template")
        
        # Update Failed count of Transaction (row 8) and Passed count of Transaction (row 9)
        # These will be calculated after processing all transactions
        
        # Clear rows 14 and 15 (unnecessary data from template)
        worksheet.delete_rows(14, 2)
        
        # Clear existing transaction data (starting from row 14, after deletion)
        max_row = worksheet.max_row
        if max_row > 13:
            worksheet.delete_rows(14, max_row - 13)
        
        # Add transaction data starting at row 14
        start_row = 14
        
        # Define styles for transaction rows
        cell_alignment = Alignment(horizontal='left', vertical='center')
        number_alignment = Alignment(horizontal='right', vertical='center')
        thin_border = Border(
            left=Side(style='thin', color='D3D3D3'),
            right=Side(style='thin', color='D3D3D3'),
            top=Side(style='thin', color='D3D3D3'),
            bottom=Side(style='thin', color='D3D3D3')
        )
        
        # Write transaction data and count pass/fail
        failed_count = 0
        passed_count = 0
        high_error_count = 0  # Count transactions with >2% error rate
        
        for i, (idx, row) in enumerate(df_transactions.iterrows()):
            row_num = start_row + i
            
            # Column A: Transaction Name
            cell = worksheet.cell(row=row_num, column=1, value=row['Element Label'])
            cell.alignment = cell_alignment
            cell.border = thin_border
            
            # Column B: Count
            cell = worksheet.cell(row=row_num, column=2, value=int(row['# Samples']))
            cell.alignment = number_alignment
            cell.border = thin_border
            
            # Column C: Avg
            cell = worksheet.cell(row=row_num, column=3, value=int(row['Avg. Response Time (ms)']))
            cell.alignment = number_alignment
            cell.border = thin_border
            
            # Column D: P95%
            cell = worksheet.cell(row=row_num, column=4, value=int(row['95% line (ms)']))
            cell.alignment = number_alignment
            cell.border = thin_border
            
            # Column E: Error% with color coding
            error_pct = row['Error Percentage']
            cell = worksheet.cell(row=row_num, column=5, value=error_pct)
            cell.alignment = number_alignment
            cell.border = thin_border
            
            # Apply color based on error percentage
            if error_pct > 2:
                cell.font = Font(color='FF0000', bold=True)  # Red for >2%
                high_error_count += 1
            elif error_pct > 1:
                cell.font = Font(color='FFA500', bold=True)  # Orange for 1-2%
            
            # Column F: Result (pass/Fail based on test type and thresholds)
            p95_value = int(row['95% line (ms)'])
            error_pct = row['Error Percentage']
            
            if test_type.upper() == 'API':
                # API: P95% > 500ms = Fail, otherwise Pass
                if p95_value > 500:
                    result = 'Fail'
                    result_color = 'FF0000'  # Red
                    failed_count += 1
                else:
                    result = 'Pass'
                    result_color = '0000FF'  # Blue
                    passed_count += 1
            else:
                # UI: P95% > 2000ms = Fail, otherwise Pass
                if p95_value > 2000:
                    result = 'Fail'
                    result_color = 'FF0000'  # Red
                    failed_count += 1
                else:
                    result = 'Pass'
                    result_color = '0000FF'  # Blue
                    passed_count += 1
            
            cell = worksheet.cell(row=row_num, column=6, value=result)
            cell.alignment = cell_alignment
            cell.border = thin_border
            cell.font = Font(color=result_color, bold=True)
        
        # Add ALL row at the end
        all_row_num = start_row + len(df_transactions)
        worksheet.cell(row=all_row_num, column=1, value='ALL')
        worksheet.cell(row=all_row_num, column=2, value=total_hits)
        worksheet.cell(row=all_row_num, column=3, value=int(avg_response_time))
        worksheet.cell(row=all_row_num, column=4, value=response_95)
        error_cell = worksheet.cell(row=all_row_num, column=5, value=round(errors, 2))
        error_cell.number_format = '0.00'  # Format as decimal number with 2 decimal places
        # Determine ALL row result
        if test_type.upper() == 'API':
            all_result = 'Fail' if response_95 > 500 else 'Pass'
            all_result_color = 'FF0000' if response_95 > 500 else '0000FF'
        else:
            all_result = 'Fail' if response_95 > 2000 else 'Pass'
            all_result_color = 'FF0000' if response_95 > 2000 else '0000FF'
        
        worksheet.cell(row=all_row_num, column=6, value=all_result)
        
        # Apply formatting to ALL row
        for col in range(1, 7):
            cell = worksheet.cell(row=all_row_num, column=col)
            cell.border = thin_border
            if col == 6:
                cell.font = Font(color=all_result_color, bold=True)
                cell.alignment = cell_alignment  # Result column should be left-aligned
            elif col > 1:
                cell.alignment = number_alignment
            else:
                cell.alignment = cell_alignment
        
        # Update row 8: Failed count of Transaction
        worksheet['B8'] = failed_count
        worksheet.merge_cells('B8:F8')
        
        # Update row 9: Passed count of Transaction
        worksheet['B9'] = passed_count
        worksheet.merge_cells('B9:F9')
        
        # Update row 10: Result
        # Fail if any of these conditions are true:
        # 1. Error rate > 1%
        # 2. P95% > threshold (500ms for API, 2000ms for UI)
        # 3. Any failed transactions exist
        
        error_rate_fail = errors > 1  # errors is already in percentage format (e.g., 0.46 = 0.46%)
        
        if test_type.upper() == 'API':
            p95_fail = response_95 > 500
        else:
            p95_fail = response_95 > 2000
        
        any_transaction_failed = failed_count > 0
        
        # Overall result is Fail if ANY condition is true
        overall_result = 'Fail' if (error_rate_fail or p95_fail or any_transaction_failed) else 'Pass'
        overall_result_color = 'FF0000' if (error_rate_fail or p95_fail or any_transaction_failed) else '0000FF'
        
        worksheet['B10'] = overall_result
        worksheet['B10'].font = Font(color=overall_result_color, bold=True)
        worksheet.merge_cells('B10:F10')
        
        # Update row 11: Analysis - explain why test failed or passed
        analysis_parts = []
        
        if error_rate_fail:
            analysis_parts.append(f"Error rate is high ({errors:.2f}%).")
        
        if any_transaction_failed:
            analysis_parts.append(f"There are {failed_count} failed transaction(s).")
        
        if high_error_count > 0:
            analysis_parts.append(f"{high_error_count} transaction(s) have more than 2% error rate.")
        
        if p95_fail:
            sla_threshold = 500 if test_type.upper() == 'API' else 2000
            analysis_parts.append(f"Overall P95% response time is over SLA ({response_95} ms > {sla_threshold} ms).")
        
        if analysis_parts:
            # Join with newline character for multi-line display in Excel cell
            analysis_text = "\n".join(analysis_parts)
        else:
            analysis_text = "All metrics are within acceptable thresholds."
        
        # Write analysis to row 11 with wrap text enabled
        worksheet['B11'] = analysis_text
        worksheet['B11'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        worksheet.merge_cells('B11:F11')
        
        # Set row 11 height based on number of lines in analysis (approximately 15 points per line)
        num_lines = analysis_text.count('\n') + 1
        worksheet.row_dimensions[11].height = num_lines * 15
        
        # Row 12: System Stats - keep empty but merge cells
        worksheet['B12'] = ''
        worksheet.merge_cells('B12:F12')
        
        # Row 13: Transaction Name header row - set height to match other rows
        worksheet.row_dimensions[13].height = 15  # Standard row height
        
        print(f"  ✓ Added {len(df_transactions)} transactions to report")
        print(f"  ✓ Failed transactions: {failed_count}")
        print(f"  ✓ Passed transactions: {passed_count}")
        print(f"  ✓ Overall result: {overall_result}")
        print(f"  ✓ Analysis: {analysis_text}")
        
    else:
        # Create new workbook without template
        print(f"  ⚠ Template not found, creating report from scratch")
        
        # Create a new workbook
        from openpyxl import Workbook
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = 'Performance Test Report'
        
        # Get test name
        test_name = os.path.splitext(os.path.basename(input_csv))[0]
        
        # Define styles
        header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        label_font = Font(name='Calibri', size=11, bold=True)
        cell_alignment = Alignment(horizontal='left', vertical='center')
        number_alignment = Alignment(horizontal='right', vertical='center')
        center_alignment = Alignment(horizontal='center', vertical='center')
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        
        # Define blue bold font for data values
        data_font = Font(name='Calibri', size=11, bold=True, color='0000FF')
        
        # Row 1: Test Name
        worksheet['A1'] = 'Test Name'
        worksheet['A1'].font = label_font
        worksheet['A1'].border = thin_border
        worksheet['B1'] = test_name
        worksheet['B1'].font = data_font
        worksheet['B1'].alignment = center_alignment
        worksheet.merge_cells('B1:F1')
        for col in range(2, 7):
            cell = worksheet.cell(row=1, column=col)
            cell.border = thin_border
            cell.font = data_font
            cell.alignment = center_alignment
        
        # Row 2: Date
        worksheet['A2'] = 'Date'
        worksheet['A2'].font = label_font
        worksheet['A2'].border = thin_border
        worksheet['B2'] = ''
        worksheet['B2'].font = data_font
        worksheet.merge_cells('B2:F2')
        for col in range(2, 7):
            cell = worksheet.cell(row=2, column=col)
            cell.border = thin_border
            cell.font = data_font
        
        # Row 3: Average Throughput
        worksheet['A3'] = 'Average Throughput (Kbytes/second)'
        worksheet['A3'].font = label_font
        worksheet['A3'].border = thin_border
        worksheet['B3'] = avg_bandwidth
        worksheet['B3'].font = data_font
        worksheet['B3'].alignment = center_alignment
        worksheet.merge_cells('B3:F3')
        for col in range(2, 7):
            cell = worksheet.cell(row=3, column=col)
            cell.border = thin_border
            cell.font = data_font
            cell.alignment = center_alignment
        
        # Row 4: Total Hits
        worksheet['A4'] = 'Total Hits'
        worksheet['A4'].font = label_font
        worksheet['A4'].border = thin_border
        worksheet['B4'] = total_hits
        worksheet['B4'].font = data_font
        worksheet['B4'].alignment = center_alignment
        worksheet.merge_cells('B4:F4')
        for col in range(2, 7):
            cell = worksheet.cell(row=4, column=col)
            cell.border = thin_border
            cell.font = data_font
            cell.alignment = center_alignment
        
        # Row 5: Average Hits per Second
        worksheet['A5'] = 'Average Hits per Second(TPS)'
        worksheet['A5'].font = label_font
        worksheet['A5'].border = thin_border
        worksheet['B5'] = avg_hits_per_sec
        worksheet['B5'].font = data_font
        worksheet['B5'].alignment = center_alignment
        worksheet.merge_cells('B5:F5')
        for col in range(2, 7):
            cell = worksheet.cell(row=5, column=col)
            cell.border = thin_border
            cell.font = data_font
            cell.alignment = center_alignment
        
        # Row 6: 95% Response Time
        worksheet['A6'] = '95% Response Time(ms)'
        worksheet['A6'].font = label_font
        worksheet['A6'].border = thin_border
        worksheet['B6'] = response_95
        worksheet['B6'].font = data_font
        worksheet['B6'].alignment = center_alignment
        worksheet.merge_cells('B6:F6')
        for col in range(2, 7):
            cell = worksheet.cell(row=6, column=col)
            cell.border = thin_border
            cell.font = data_font
            cell.alignment = center_alignment
        
        # Row 7: Errors
        worksheet['A7'] = 'Errors'
        worksheet['A7'].font = label_font
        worksheet['A7'].border = thin_border
        worksheet['B7'] = f"{errors:.2f}%"
        worksheet['B7'].font = data_font
        worksheet['B7'].alignment = center_alignment
        worksheet.merge_cells('B7:F7')
        for col in range(2, 7):
            cell = worksheet.cell(row=7, column=col)
            cell.border = thin_border
            cell.font = data_font
            cell.alignment = center_alignment
        
        # Count pass/fail transactions
        failed_count = 0
        passed_count = 0
        high_error_count = 0
        
        for idx, row in df_transactions.iterrows():
            p95_value = int(row['95% line (ms)'])
            error_pct = row['Error Percentage']
            
            if test_type.upper() == 'API':
                if p95_value > 500:
                    failed_count += 1
                else:
                    passed_count += 1
            else:
                if p95_value > 2000:
                    failed_count += 1
                else:
                    passed_count += 1
            
            if error_pct > 2:
                high_error_count += 1
        
        # Row 8: Failed count of Transaction
        worksheet['A8'] = 'Failed count of Transaction'
        worksheet['A8'].font = label_font
        worksheet['A8'].border = thin_border
        worksheet['B8'] = failed_count
        worksheet['B8'].font = data_font
        worksheet['B8'].alignment = center_alignment
        worksheet.merge_cells('B8:F8')
        for col in range(2, 7):
            cell = worksheet.cell(row=8, column=col)
            cell.border = thin_border
            cell.font = data_font
            cell.alignment = center_alignment
        
        # Row 9: Passed count of Transaction
        worksheet['A9'] = 'Passed count of Transaction'
        worksheet['A9'].font = label_font
        worksheet['A9'].border = thin_border
        worksheet['B9'] = passed_count
        worksheet['B9'].font = data_font
        worksheet['B9'].alignment = center_alignment
        worksheet.merge_cells('B9:F9')
        for col in range(2, 7):
            cell = worksheet.cell(row=9, column=col)
            cell.border = thin_border
            cell.font = data_font
            cell.alignment = center_alignment
        
        # Calculate overall result
        error_rate_fail = errors > 1
        if test_type.upper() == 'API':
            p95_fail = response_95 > 500
        else:
            p95_fail = response_95 > 2000
        any_transaction_failed = failed_count > 0
        
        overall_result = 'Fail' if (error_rate_fail or p95_fail or any_transaction_failed) else 'Pass'
        overall_result_color = 'FF0000' if (error_rate_fail or p95_fail or any_transaction_failed) else '0000FF'
        
        # Row 10: Result
        worksheet['A10'] = 'Result'
        worksheet['A10'].font = label_font
        worksheet['A10'].border = thin_border
        worksheet['B10'] = overall_result
        worksheet['B10'].font = Font(color=overall_result_color, bold=True)
        worksheet['B10'].alignment = center_alignment
        worksheet.merge_cells('B10:F10')
        for col in range(2, 7):
            cell = worksheet.cell(row=10, column=col)
            cell.border = thin_border
            cell.font = Font(color=overall_result_color, bold=True)
            cell.alignment = center_alignment
        
        # Row 11: Analysis
        analysis_parts = []
        if error_rate_fail:
            analysis_parts.append(f"Error rate is high ({errors:.2f}%).")
        if any_transaction_failed:
            analysis_parts.append(f"There are {failed_count} failed transaction(s).")
        if high_error_count > 0:
            analysis_parts.append(f"{high_error_count} transaction(s) have more than 2% error rate.")
        if p95_fail:
            sla_threshold = 500 if test_type.upper() == 'API' else 2000
            analysis_parts.append(f"Overall P95% response time is over SLA ({response_95} ms > {sla_threshold} ms).")
        
        if analysis_parts:
            analysis_text = "\n".join(analysis_parts)
        else:
            analysis_text = "All metrics are within acceptable thresholds."
        
        worksheet['A11'] = 'Analysis'
        worksheet['A11'].font = label_font
        worksheet['A11'].border = thin_border
        worksheet['B11'] = analysis_text
        worksheet['B11'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        worksheet.merge_cells('B11:F11')
        for col in range(2, 7):
            worksheet.cell(row=11, column=col).border = thin_border
        num_lines = analysis_text.count('\n') + 1
        worksheet.row_dimensions[11].height = num_lines * 15
        
        # Row 12: System Stats
        worksheet['A12'] = 'System Stats'
        worksheet['A12'].font = label_font
        worksheet['A12'].border = thin_border
        worksheet['B12'] = ''
        worksheet.merge_cells('B12:F12')
        for col in range(2, 7):
            worksheet.cell(row=12, column=col).border = thin_border
        
        # Row 13: Transaction headers (removed empty rows 13 and 14)
        headers = ['Transaction Name', 'Count', 'Avg', 'P95%', 'Error%', 'Result']
        for col_idx, header in enumerate(headers, 1):
            cell = worksheet.cell(row=13, column=col_idx, value=header)
            cell.font = label_font  # Use label_font instead of header_font (no white text)
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # Add transaction data starting at row 14
        start_row = 14
        
        for i, (idx, row) in enumerate(df_transactions.iterrows()):
            row_num = start_row + i
            
            # Column A: Transaction Name
            cell = worksheet.cell(row=row_num, column=1, value=row['Element Label'])
            cell.alignment = cell_alignment
            cell.border = thin_border
            
            # Column B: Count
            cell = worksheet.cell(row=row_num, column=2, value=int(row['# Samples']))
            cell.alignment = number_alignment
            cell.border = thin_border
            
            # Column C: Avg
            cell = worksheet.cell(row=row_num, column=3, value=int(row['Avg. Response Time (ms)']))
            cell.alignment = number_alignment
            cell.border = thin_border
            
            # Column D: P95%
            cell = worksheet.cell(row=row_num, column=4, value=int(row['95% line (ms)']))
            cell.alignment = number_alignment
            cell.border = thin_border
            
            # Column E: Error% with color coding
            error_pct = row['Error Percentage']
            cell = worksheet.cell(row=row_num, column=5, value=error_pct)
            cell.alignment = number_alignment
            cell.border = thin_border
            
            if error_pct > 2:
                cell.font = Font(color='FF0000', bold=True)
            elif error_pct > 1:
                cell.font = Font(color='FFA500', bold=True)
            
            # Column F: Result
            p95_value = int(row['95% line (ms)'])
            
            if test_type.upper() == 'API':
                if p95_value > 500:
                    result = 'Fail'
                    result_color = 'FF0000'
                else:
                    result = 'Pass'
                    result_color = '0000FF'
            else:
                if p95_value > 2000:
                    result = 'Fail'
                    result_color = 'FF0000'
                else:
                    result = 'Pass'
                    result_color = '0000FF'
            
            cell = worksheet.cell(row=row_num, column=6, value=result)
            cell.alignment = center_alignment
            cell.border = thin_border
            cell.font = Font(color=result_color, bold=True)
        
        # Add ALL row at the end
        all_row_num = start_row + len(df_transactions)
        worksheet.cell(row=all_row_num, column=1, value='ALL')
        worksheet.cell(row=all_row_num, column=2, value=total_hits)
        worksheet.cell(row=all_row_num, column=3, value=int(avg_response_time))
        worksheet.cell(row=all_row_num, column=4, value=response_95)
        error_cell = worksheet.cell(row=all_row_num, column=5, value=round(errors, 2))
        error_cell.number_format = '0.00'
        
        # Determine ALL row result
        if test_type.upper() == 'API':
            all_result = 'Fail' if response_95 > 500 else 'Pass'
            all_result_color = 'FF0000' if response_95 > 500 else '0000FF'
        else:
            all_result = 'Fail' if response_95 > 2000 else 'Pass'
            all_result_color = 'FF0000' if response_95 > 2000 else '0000FF'
        
        worksheet.cell(row=all_row_num, column=6, value=all_result)
        
        # Apply formatting to ALL row
        for col in range(1, 7):
            cell = worksheet.cell(row=all_row_num, column=col)
            cell.border = thin_border
            if col == 6:
                cell.font = Font(color=all_result_color, bold=True)
                cell.alignment = center_alignment
            elif col > 1:
                cell.alignment = number_alignment
            else:
                cell.alignment = cell_alignment
        
        # Set column widths
        worksheet.column_dimensions['A'].width = 35
        worksheet.column_dimensions['B'].width = 12
        worksheet.column_dimensions['C'].width = 12
        worksheet.column_dimensions['D'].width = 12
        worksheet.column_dimensions['E'].width = 12
        worksheet.column_dimensions['F'].width = 12
        
        print(f"  ✓ Created report from scratch")
        print(f"  ✓ Added {len(df_transactions)} transactions to report")
        print(f"  ✓ Failed transactions: {failed_count}")
        print(f"  ✓ Passed transactions: {passed_count}")
        print(f"  ✓ Overall result: {overall_result}")
        print(f"  ✓ Analysis: {analysis_text}")
    
    # Save the workbook
    workbook.save(output_xlsx)
    print(f"  ✓ Successfully created: {output_xlsx}")
    
    # Print summary
    print(f"\n  Summary:")
    print(f"    - Total transactions: {len(df_transactions)}")
    print(f"    - Total hits: {total_hits:,}")
    print(f"    - Average throughput: {avg_bandwidth} KB/s")
    print(f"    - Error rate: {errors:.2f}%")
    print(f"    - Output file size: {os.path.getsize(output_xlsx):,} bytes")
    
    return output_xlsx


def main():
    """Main function to handle command-line execution."""
    if len(sys.argv) < 3:
        print(__doc__)
        print("\nError: Please provide an input CSV file and test type (API or UI).")
        print("\nUsage:")
        print("  python blazemeter_to_excel.py <input_csv_file> <test_type> [output_xlsx_file]")
        print("\nTest Type:")
        print("  API - P95% > 500ms = Fail (red), P95% <= 500ms = Pass (blue)")
        print("  UI  - P95% > 2000ms = Fail (red), P95% <= 2000ms = Pass (blue)")
        print("\nWildcard Support:")
        print("  You can use wildcards to process multiple files:")
        print("  python blazemeter_to_excel.py 'TrueFlix*.csv' UI")
        print("  python blazemeter_to_excel.py '*.csv' API")
        sys.exit(1)
    
    input_pattern = sys.argv[1]
    test_type = sys.argv[2]
    output_xlsx = sys.argv[3] if len(sys.argv) > 3 else None
    
    # Validate test type
    if test_type.upper() not in ['API', 'UI']:
        print(f"\n✗ Error: Invalid test type '{test_type}'. Must be 'API' or 'UI'.")
        sys.exit(1)
    
    # Check if input pattern contains wildcards
    if '*' in input_pattern or '?' in input_pattern:
        # Find all matching files
        csv_files = glob.glob(input_pattern)
        
        if not csv_files:
            print(f"\n✗ Error: No files found matching pattern '{input_pattern}'")
            sys.exit(1)
        
        # Sort files alphabetically
        csv_files.sort()
        
        print(f"\n{'='*80}")
        print(f"Found {len(csv_files)} file(s) matching pattern '{input_pattern}'")
        print(f"{'='*80}")
        
        # Process each file
        success_count = 0
        failed_files = []
        
        for i, csv_file in enumerate(csv_files, 1):
            print(f"\n[{i}/{len(csv_files)}] Processing: {csv_file}")
            print(f"{'-'*80}")
            
            try:
                convert_blazemeter_to_excel(csv_file, test_type, None)
                success_count += 1
            except Exception as e:
                print(f"\n✗ Error processing {csv_file}: {e}")
                failed_files.append(csv_file)
                import traceback
                traceback.print_exc()
        
        # Print final summary
        print(f"\n{'='*80}")
        print(f"BATCH PROCESSING COMPLETE")
        print(f"{'='*80}")
        print(f"  ✓ Successfully processed: {success_count}/{len(csv_files)} files")
        
        if failed_files:
            print(f"  ✗ Failed files ({len(failed_files)}):")
            for failed_file in failed_files:
                print(f"    - {failed_file}")
            sys.exit(1)
        else:
            print(f"  All files processed successfully!")
    else:
        # Single file processing
        if output_xlsx and '*' in output_xlsx:
            print(f"\n✗ Error: Output filename cannot contain wildcards")
            sys.exit(1)
        
        try:
            convert_blazemeter_to_excel(input_pattern, test_type, output_xlsx)
        except Exception as e:
            print(f"\n✗ Error: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)


if __name__ == "__main__":
    main()
