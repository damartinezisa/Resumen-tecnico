#!/usr/bin/env python3
"""
Script to extract tables from PDF to Excel without using OCR.
This provides a two-step process:
1. Extract PDF tables to Excel for pre-validation
2. Convert validated Excel to JSON
"""

import pdfplumber
import pandas as pd
import json
import sys
import os
from typing import List, Dict, Any
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment


def extract_pdf_to_excel(pdf_path: str, page_number: int = 2, output_excel: str = None):
    """
    Extract tables from PDF page and save to Excel for validation.
    
    Args:
        pdf_path: Path to the PDF file
        page_number: Page number to extract (1-indexed)
        output_excel: Output Excel file path (default: based on PDF name)
    
    Returns:
        Path to the created Excel file
    """
    if output_excel is None:
        base_name = os.path.splitext(pdf_path)[0]
        output_excel = f"{base_name}_page{page_number}_tables.xlsx"
    
    print(f"Extracting tables from {pdf_path}, page {page_number}...", file=sys.stderr)
    
    with pdfplumber.open(pdf_path) as pdf:
        if page_number > len(pdf.pages):
            raise ValueError(f"PDF only has {len(pdf.pages)} pages, cannot extract page {page_number}")
        
        # Get the specified page (1-indexed)
        page = pdf.pages[page_number - 1]
        
        # Extract tables using line-based detection
        tables = page.extract_tables({
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "snap_tolerance": 3,
            "join_tolerance": 3,
        })
        
        print(f"Found {len(tables)} table(s)", file=sys.stderr)
        
        # Create Excel writer
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            if not tables:
                print("Warning: No tables found!", file=sys.stderr)
                # Create empty sheet
                df = pd.DataFrame(["No tables found on this page"])
                df.to_excel(writer, sheet_name='Info', index=False, header=False)
            else:
                # Process the main table (should contain all 3 tables merged)
                main_table = tables[0]
                
                # Create a DataFrame from the raw table
                df = pd.DataFrame(main_table)
                
                # Save raw extraction to first sheet
                df.to_excel(writer, sheet_name='Raw_Data', index=False, header=False)
                
                # Now parse and separate the three logical tables
                # Based on the structure, we need to identify:
                # 1. ANALISIS table (rows 0-32)
                # 2. PRODUCTO TERMINADO (AZUCAR) table (rows 33-37)
                # 3. PRODUCTO TERMINADO (AZUCAR) - CONTINUACIÓN table (rows 38+)
                
                analisis_data = []
                producto_data = []
                continuacion_data = []
                
                # Find section boundaries
                producto_start = None
                continuacion_start = None
                
                for i, row in enumerate(main_table):
                    first_col = str(row[0]) if row[0] else ""
                    first_col_clean = first_col.replace("*", "").replace(" ", "").upper()
                    
                    if "PRODUCTO" in first_col_clean and "TERMINADO" in first_col_clean:
                        producto_start = i
                    elif "DESCRIPCION" in first_col.upper() and i > 30 and "Quintales" in " ".join([str(c) for c in row if c]):
                        continuacion_start = i
                        break
                
                # Split data into sections
                if producto_start and continuacion_start:
                    analisis_data = main_table[:producto_start]
                    producto_data = main_table[producto_start:continuacion_start]
                    continuacion_data = main_table[continuacion_start:]
                elif producto_start:
                    analisis_data = main_table[:producto_start]
                    producto_data = main_table[producto_start:]
                else:
                    analisis_data = main_table
                
                # Save separated tables to individual sheets
                if analisis_data:
                    df_analisis = pd.DataFrame(analisis_data)
                    df_analisis.to_excel(writer, sheet_name='ANALISIS', index=False, header=False)
                
                if producto_data:
                    df_producto = pd.DataFrame(producto_data)
                    df_producto.to_excel(writer, sheet_name='PRODUCTO_TERMINADO', index=False, header=False)
                
                if continuacion_data:
                    df_continuacion = pd.DataFrame(continuacion_data)
                    df_continuacion.to_excel(writer, sheet_name='CONTINUACION', index=False, header=False)
        
        # Apply formatting to make validation easier
        wb = openpyxl.load_workbook(output_excel)
        
        # Format each sheet
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # Auto-adjust column widths
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width
            
            # Highlight header rows (yellow background)
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            for row in ws.iter_rows(min_row=1, max_row=3):
                for cell in row:
                    if cell.value and any(keyword in str(cell.value).upper() for keyword in ['ANALISIS', 'HOY', 'HASTA', 'BRIX', 'PRODUCTO', 'DESCRIPCION']):
                        cell.fill = yellow_fill
                        cell.font = Font(bold=True)
        
        wb.save(output_excel)
        
    print(f"✓ Excel file created: {output_excel}", file=sys.stderr)
    print(f"✓ Please validate the data in Excel and save any corrections.", file=sys.stderr)
    print(f"✓ Then use 'python {sys.argv[0]} --excel-to-json {output_excel}' to convert to JSON", file=sys.stderr)
    
    return output_excel


def excel_to_json(excel_path: str, output_json: str = None):
    """
    Convert validated Excel file to JSON format.
    
    Args:
        excel_path: Path to the Excel file
        output_json: Output JSON file path (default: based on Excel name)
    
    Returns:
        Path to the created JSON file
    """
    if output_json is None:
        base_name = os.path.splitext(excel_path)[0]
        output_json = f"{base_name}.json"
    
    print(f"Converting Excel to JSON: {excel_path}...", file=sys.stderr)
    
    # Read Excel file
    xl_file = pd.ExcelFile(excel_path)
    
    results = []
    
    # Check which sheets exist
    has_analisis = 'ANALISIS' in xl_file.sheet_names
    has_producto = 'PRODUCTO_TERMINADO' in xl_file.sheet_names
    has_continuacion = 'CONTINUACION' in xl_file.sheet_names
    
    # Parse ANALISIS sheet
    if has_analisis:
        df_analisis = pd.read_excel(excel_path, sheet_name='ANALISIS', header=None)
        hoy_indicators, hasta_indicators = parse_analisis_from_df(df_analisis)
        
        results.append({
            "Nombre Agrupador": "ANALISIS - HOY",
            "Indicadores": hoy_indicators
        })
        results.append({
            "Nombre Agrupador": "ANALISIS - HASTA",
            "Indicadores": hasta_indicators
        })
    
    # Parse PRODUCTO TERMINADO sheet
    if has_producto:
        df_producto = pd.read_excel(excel_path, sheet_name='PRODUCTO_TERMINADO', header=None)
        producto_indicators = parse_producto_from_df(df_producto)
        
        results.append({
            "Nombre Agrupador": "PRODUCTO TERMINADO (AZUCAR)",
            "Indicadores": producto_indicators
        })
    
    # Parse CONTINUACION sheet
    if has_continuacion:
        df_continuacion = pd.read_excel(excel_path, sheet_name='CONTINUACION', header=None)
        continuacion_indicators = parse_continuacion_from_df(df_continuacion)
        
        results.append({
            "Nombre Agrupador": "PRODUCTO TERMINADO (AZUCAR) - CONTINUACIÓN",
            "Indicadores": continuacion_indicators
        })
    
    # Write JSON
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=4, ensure_ascii=False)
    
    print(f"✓ JSON file created: {output_json}", file=sys.stderr)
    
    return output_json


def parse_analisis_from_df(df):
    """Parse ANALISIS table from DataFrame"""
    hoy_indicators = []
    hasta_indicators = []
    
    # The ANALISIS section is from row 0 to before PRODUCTO TERMINADO
    # Find where PRODUCTO TERMINADO starts
    producto_start = None
    for row_idx in range(len(df)):
        row = df.iloc[row_idx]
        first_val = str(row[0]).upper() if pd.notna(row[0]) else ""
        if "PRODUCTO" in first_val and "TERMINADO" in first_val:
            producto_start = row_idx
            break
    
    if producto_start is None:
        producto_start = len(df)
    
    # Process ANALISIS rows (row 2 onwards until PRODUCTO starts)
    for row_idx in range(2, producto_start):
        row = df.iloc[row_idx]
        
        # First column is DESCRIPCION
        descripcion = str(row[0]).strip() if pd.notna(row[0]) else ""
        
        if not descripcion or descripcion == "nan" or len(descripcion) < 2:
            continue
        
        # Skip separator rows and empty rows
        if all(c in '—-_=*\t ' for c in descripcion):
            continue
        
        # Create HOY indicator
        # HOY columns: BRIX(3), SAC(4), PZA(7), COLOR(16), PH(18), GR(20)
        hoy_indicator = {
            "DESCRIPCION": descripcion,
            "BRIX": parse_value(row[3]),
            "SAC": parse_value(row[4]),
            "PZA": parse_value(row[7]),
            "COLOR": parse_value(row[16]),
            "PH": parse_value(row[18]),
            "GR": parse_value(row[20])
        }
        
        # Create HASTA indicator
        # HASTA columns: BRIX(9), SAC(11), PZA(14)
        hasta_indicator = {
            "DESCRIPCION": descripcion,
            "BRIX": parse_value(row[9]),
            "SAC": parse_value(row[11]),
            "PZA": parse_value(row[14])
        }
        
        hoy_indicators.append(hoy_indicator)
        hasta_indicators.append(hasta_indicator)
    
    return hoy_indicators, hasta_indicators


def parse_producto_from_df(df):
    """Parse PRODUCTO TERMINADO table from DataFrame"""
    indicators = []
    
    # The PRODUCTO sheet has been separated, so rows are:
    # Row 0: Section title
    # Row 1: Column headers (partial)
    # Row 2: ESTANDAR/DIA values
    # Row 3: HOY values
    # Row 4: HASTA values
    
    if len(df) < 5:
        return indicators
    
    # Find the ESTANDAR, HOY, HASTA rows
    estandar_row_idx = None
    hoy_row_idx = None
    hasta_row_idx = None
    
    for row_idx in range(len(df)):
        first_val = str(df.iloc[row_idx, 0]).strip().upper() if pd.notna(df.iloc[row_idx, 0]) else ""
        
        if "ESTANDAR" in first_val:
            estandar_row_idx = row_idx
        elif first_val == "HOY":
            hoy_row_idx = row_idx
        elif first_val == "HASTA" and hoy_row_idx is not None:
            hasta_row_idx = row_idx
            break
    
    if estandar_row_idx is None or hoy_row_idx is None or hasta_row_idx is None:
        return indicators
    
    # Column descriptions - extract from the data columns
    # The table structure has column headers scattered across row 1
    # Columns: TOTAL QQ CRUDA (1), TOTAL QQ ESTAN. (2), TOTAL QQ REFIN. (5), QQ PRODUCIDOS (8), etc.
    
    # Based on inspection, the columns are at specific indices
    column_mapping = [
        ("TOTAL QQ CRUDA", 1),
        ("TOTAL QQ ESTAN.", 2),
        ("TOTAL QQ REFIN.", 5),
        ("QQ PRODUCIDOS", 8),
        ("AZ. EQUIV. (MIEL)", 11),
        ("AZ. PMR", 14),
        ("TOTAL QUINTALES", 17)
    ]
    
    for desc, col_idx in column_mapping:
        indicator = {
            "DESCRIPCION": desc,
            "ESTANDAR/DIA": parse_value(df.iloc[estandar_row_idx, col_idx]) if col_idx < len(df.columns) else None,
            "HOY": parse_value(df.iloc[hoy_row_idx, col_idx]) if col_idx < len(df.columns) else None,
            "HASTA": parse_value(df.iloc[hasta_row_idx, col_idx]) if col_idx < len(df.columns) else None
        }
        indicators.append(indicator)
    
    return indicators


def parse_continuacion_from_df(df):
    """Parse CONTINUACION table from DataFrame"""
    indicators = []
    
    # The CONTINUACION sheet has been separated
    # Row 0: Column headers
    # Rows 1+: Data rows
    
    if len(df) < 2:
        return indicators
    
    # Column names (fixed positions based on structure)
    column_names = ["DESCRIPCION", "Quintales", "% HUM.", "% CEN.", "% POL.", "COLOR", 
                    "Mg/kg Vit.\"A\"", "T. GRANO", "% C.V", "FS", "TEMP. ºC.", "SED."]
    
    # Process data rows (skip header row 0)
    for row_idx in range(1, len(df)):
        row = df.iloc[row_idx]
        
        descripcion = str(row[0]).strip() if pd.notna(row[0]) else ""
        
        if not descripcion or descripcion == "nan" or len(descripcion) < 2:
            continue
        
        # Skip separator rows
        if all(c in '—-_=*\t ' for c in descripcion):
            continue
        
        indicator = {"DESCRIPCION": descripcion}
        
        # Map columns based on actual positions in the sheet
        # From inspection: Quintales(2), % HUM(4), % CEN(6), % POL(8), COLOR(9), etc.
        col_mapping = {
            "Quintales": 2,
            "% HUM.": 4,
            "% CEN.": 6,
            "% POL.": 8,
            "COLOR": 9,
            "Mg/kg Vit.\"A\"": 10,
            "T. GRANO": 11,
            "% C.V": 13,
            "FS": 14,
            "TEMP. ºC.": 15,
            "SED.": 16
        }
        
        for col_name in column_names[1:]:
            col_idx = col_mapping.get(col_name)
            if col_idx is not None and col_idx < len(row):
                indicator[col_name] = parse_value(row[col_idx])
            else:
                indicator[col_name] = None
        
        indicators.append(indicator)
    
    return indicators


def parse_value(val):
    """Parse a value from Excel cell to appropriate Python type"""
    if pd.isna(val):
        return None
    
    val_str = str(val).strip()
    
    if not val_str or val_str == "" or val_str.lower() == "nan":
        return None
    
    # Remove thousands separators (commas)
    val_str = val_str.replace(',', '')
    
    # Try to parse as number
    try:
        if '.' in val_str:
            return float(val_str)
        else:
            return int(val_str)
    except ValueError:
        # Return as string if not a number
        return val_str if val_str else None


def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage:")
        print(f"  Step 1 - Extract PDF to Excel:")
        print(f"    python {sys.argv[0]} <pdf_file> [page_number]")
        print(f"")
        print(f"  Step 2 - Convert Excel to JSON:")
        print(f"    python {sys.argv[0]} --excel-to-json <excel_file> [output_json]")
        print(f"")
        print(f"Example:")
        print(f"  python {sys.argv[0]} 1003.pdf 2")
        print(f"  # Validate data in Excel, then:")
        print(f"  python {sys.argv[0]} --excel-to-json 1003_page2_tables.xlsx")
        sys.exit(1)
    
    if sys.argv[1] == "--excel-to-json":
        # Convert Excel to JSON
        if len(sys.argv) < 3:
            print("Error: Excel file path required", file=sys.stderr)
            sys.exit(1)
        
        excel_path = sys.argv[2]
        output_json = sys.argv[3] if len(sys.argv) > 3 else None
        
        json_path = excel_to_json(excel_path, output_json)
        
        # Output only JSON to stdout
        with open(json_path, 'r', encoding='utf-8') as f:
            print(f.read())
    
    else:
        # Extract PDF to Excel
        pdf_path = sys.argv[1]
        page_number = int(sys.argv[2]) if len(sys.argv) > 2 else 2
        
        excel_path = extract_pdf_to_excel(pdf_path, page_number)


if __name__ == "__main__":
    main()
