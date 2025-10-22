#!/usr/bin/env python3
"""
Script to convert validated Excel files to JSON format.
This is the second step in a two-step process:
1. Extract PDF tables to Excel for pre-validation (use pdf-to-excel.py)
2. Convert validated Excel to JSON (this script)

This script reads Excel files created by pdf-to-excel.py and converts them
to a structured JSON format suitable for further processing or API consumption.
"""

# Import required libraries for Excel reading, JSON creation, and data processing
import pandas as pd  # Library for reading Excel files and data manipulation
import json  # Library for creating and writing JSON files
import sys  # System-specific parameters and functions (for command-line arguments)
import os  # Operating system interface (for file path operations)
from typing import List, Dict, Any  # Type hints for better code documentation


def excel_to_json(excel_path: str, output_json: str = None):
    """
    Convert validated Excel file to structured JSON format.
    
    This function reads an Excel file with multiple sheets (created by pdf-to-excel.py)
    and converts the data to a structured JSON format with proper groupings and indicators.
    
    Args:
        excel_path: Path to the Excel file to convert
        output_json: Output JSON file path (default: auto-generated from Excel name)
    
    Returns:
        Path to the created JSON file
    """
    # If no output filename is specified, generate one based on the Excel name
    if output_json is None:
        base_name = os.path.splitext(excel_path)[0]  # Remove .xlsx extension
        output_json = f"{base_name}.json"
    
    # Print progress message to stderr
    print(f"Converting Excel to JSON: {excel_path}...", file=sys.stderr)
    
    # Read the Excel file to access all sheets
    xl_file = pd.ExcelFile(excel_path)
    
    # Initialize list to store all result groups
    results = []
    
    # Check which sheets exist in the Excel file
    # This allows the script to work with files that may have different sheet combinations
    has_analisis = 'ANALISIS' in xl_file.sheet_names
    has_producto = 'PRODUCTO_TERMINADO' in xl_file.sheet_names
    has_continuacion = 'CONTINUACION' in xl_file.sheet_names
    
    # Parse ANALISIS sheet if it exists
    # This sheet contains data that needs to be split into two groups: HOY and HASTA
    if has_analisis:
        # Read the ANALISIS sheet without assuming any header row
        df_analisis = pd.read_excel(excel_path, sheet_name='ANALISIS', header=None)
        
        # Parse the sheet and get two lists of indicators
        hoy_indicators, hasta_indicators = parse_analisis_from_df(df_analisis)
        
        # Add HOY group to results
        results.append({
            "Nombre Agrupador": "ANALISIS - HOY",
            "Indicadores": hoy_indicators
        })
        
        # Add HASTA group to results
        results.append({
            "Nombre Agrupador": "ANALISIS - HASTA",
            "Indicadores": hasta_indicators
        })
    
    # Parse PRODUCTO TERMINADO sheet if it exists
    # This sheet contains production data with ESTANDAR/DIA, HOY, and HASTA rows
    if has_producto:
        # Read the PRODUCTO_TERMINADO sheet without assuming any header row
        df_producto = pd.read_excel(excel_path, sheet_name='PRODUCTO_TERMINADO', header=None)
        
        # Parse the sheet and get list of indicators
        producto_indicators = parse_producto_from_df(df_producto)
        
        # Add PRODUCTO TERMINADO group to results
        results.append({
            "Nombre Agrupador": "PRODUCTO TERMINADO (AZUCAR)",
            "Indicadores": producto_indicators
        })
    
    # Parse CONTINUACION sheet if it exists
    # This sheet contains additional production data with various quality metrics
    if has_continuacion:
        # Read the CONTINUACION sheet without assuming any header row
        df_continuacion = pd.read_excel(excel_path, sheet_name='CONTINUACION', header=None)
        
        # Parse the sheet and get list of indicators
        continuacion_indicators = parse_continuacion_from_df(df_continuacion)
        
        # Add CONTINUACIÓN group to results
        results.append({
            "Nombre Agrupador": "PRODUCTO TERMINADO (AZUCAR) - CONTINUACIÓN",
            "Indicadores": continuacion_indicators
        })
    
    # Write the results to a JSON file with proper formatting
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(
            results,
            f,
            indent=4,  # Use 4 spaces for indentation (readable formatting)
            ensure_ascii=False  # Allow Unicode characters (e.g., Spanish accents)
        )
    
    # Print success message
    print(f"✓ JSON file created: {output_json}", file=sys.stderr)
    
    return output_json


def parse_analisis_from_df(df):
    """
    Parse ANALISIS table from DataFrame and split into HOY and HASTA indicators.
    
    The ANALISIS table contains data with two main column groups:
    - HOY (today): Current measurements
    - HASTA (accumulated): Cumulative measurements
    
    Args:
        df: DataFrame containing the ANALISIS table data
    
    Returns:
        Tuple of (hoy_indicators, hasta_indicators) - two lists of indicator dictionaries
    """
    # Initialize lists for the two groups
    hoy_indicators = []
    hasta_indicators = []
    
    # Find where PRODUCTO TERMINADO section starts (end of ANALISIS section)
    producto_start = None
    for row_idx in range(len(df)):
        row = df.iloc[row_idx]
        first_val = str(row[0]).upper() if pd.notna(row[0]) else ""
        
        # Look for "PRODUCTO TERMINADO" keyword
        if "PRODUCTO" in first_val and "TERMINADO" in first_val:
            producto_start = row_idx
            break
    
    # If not found, assume ANALISIS goes to the end of the DataFrame
    if producto_start is None:
        producto_start = len(df)
    
    # Process ANALISIS data rows (starting from row 2 to skip headers, until PRODUCTO section)
    for row_idx in range(2, producto_start):
        row = df.iloc[row_idx]
        
        # Get the description from the first column
        descripcion = str(row[0]).strip() if pd.notna(row[0]) else ""
        
        # Skip empty rows or invalid descriptions
        if not descripcion or descripcion == "nan" or len(descripcion) < 2:
            continue
        
        # Skip separator rows (rows with only dashes, underscores, etc.)
        if all(c in '—-_=*\t ' for c in descripcion):
            continue
        
        # Create HOY (today) indicator with measurements from specific columns
        # Column positions based on the PDF table structure:
        # BRIX(3), SAC(4), PZA(7), COLOR(16), PH(18), GR(20)
        hoy_indicator = {
            "DESCRIPCION": descripcion,
            "BRIX": parse_value(row[3]),    # Brix measurement (sugar content)
            "SAC": parse_value(row[4]),     # SAC measurement
            "PZA": parse_value(row[7]),     # PZA measurement (purity)
            "COLOR": parse_value(row[16]),  # Color measurement
            "PH": parse_value(row[18]),     # pH measurement
            "GR": parse_value(row[20])      # GR measurement
        }
        
        # Create HASTA (accumulated) indicator with measurements from specific columns
        # Column positions: BRIX(9), SAC(11), PZA(14)
        hasta_indicator = {
            "DESCRIPCION": descripcion,
            "BRIX": parse_value(row[9]),    # Accumulated Brix
            "SAC": parse_value(row[11]),    # Accumulated SAC
            "PZA": parse_value(row[14])     # Accumulated PZA
        }
        
        # Add both indicators to their respective lists
        hoy_indicators.append(hoy_indicator)
        hasta_indicators.append(hasta_indicator)
    
    return hoy_indicators, hasta_indicators


def parse_producto_from_df(df):
    """
    Parse PRODUCTO TERMINADO table from DataFrame.
    
    This table contains production data organized in rows for:
    - ESTANDAR/DIA: Standard daily production targets
    - HOY: Today's production
    - HASTA: Accumulated production
    
    Args:
        df: DataFrame containing the PRODUCTO_TERMINADO table data
    
    Returns:
        List of indicator dictionaries, one for each column metric
    """
    indicators = []
    
    # Validate minimum row count
    if len(df) < 5:
        return indicators
    
    # Find the row indices for ESTANDAR, HOY, and HASTA data
    estandar_row_idx = None
    hoy_row_idx = None
    hasta_row_idx = None
    
    # Scan rows to find the three data rows
    for row_idx in range(len(df)):
        first_val = str(df.iloc[row_idx, 0]).strip().upper() if pd.notna(df.iloc[row_idx, 0]) else ""
        
        # Identify each row by its first column value
        if "ESTANDAR" in first_val:
            estandar_row_idx = row_idx
        elif first_val == "HOY":
            hoy_row_idx = row_idx
        elif first_val == "HASTA" and hoy_row_idx is not None:
            hasta_row_idx = row_idx
            break  # Found all three rows
    
    # If any row is missing, return empty list
    if estandar_row_idx is None or hoy_row_idx is None or hasta_row_idx is None:
        return indicators
    
    # Define column mappings: (description, column_index)
    # These indices correspond to the positions of different metrics in the table
    column_mapping = [
        ("TOTAL QQ CRUDA", 1),           # Total quintals of raw sugar
        ("TOTAL QQ ESTAN.", 2),          # Total quintals standard
        ("TOTAL QQ REFIN.", 5),          # Total quintals refined
        ("QQ PRODUCIDOS", 8),            # Quintals produced
        ("AZ. EQUIV. (MIEL)", 11),       # Sugar equivalent (molasses)
        ("AZ. PMR", 14),                 # PMR sugar
        ("TOTAL QUINTALES", 17)          # Total quintals
    ]
    
    # Create an indicator for each column metric
    for desc, col_idx in column_mapping:
        indicator = {
            "DESCRIPCION": desc,
            # Extract values from the three data rows for this column
            "ESTANDAR/DIA": parse_value(df.iloc[estandar_row_idx, col_idx]) if col_idx < len(df.columns) else None,
            "HOY": parse_value(df.iloc[hoy_row_idx, col_idx]) if col_idx < len(df.columns) else None,
            "HASTA": parse_value(df.iloc[hasta_row_idx, col_idx]) if col_idx < len(df.columns) else None
        }
        indicators.append(indicator)
    
    return indicators


def parse_continuacion_from_df(df):
    """
    Parse CONTINUACION table from DataFrame.
    
    This table contains additional quality and production metrics for finished products.
    Each row represents a different product or batch with various measurements.
    
    Args:
        df: DataFrame containing the CONTINUACION table data
    
    Returns:
        List of indicator dictionaries, one for each data row
    """
    indicators = []
    
    # Validate minimum row count
    if len(df) < 2:
        return indicators
    
    # Define column mappings: (metric_name, column_index)
    # These indices correspond to the positions of different quality metrics
    col_mapping = {
        "Quintales": 2,           # Quantity in quintals
        "% HUM.": 4,              # Humidity percentage
        "% CEN.": 6,              # Ash percentage
        "% POL.": 8,              # Polarization percentage
        "COLOR": 9,               # Color measurement
        "Mg/kg Vit.\"A\"": 10,    # Vitamin A content
        "T. GRANO": 11,           # Grain size
        "% C.V": 13,              # Coefficient of variation percentage
        "FS": 14,                 # FS measurement
        "TEMP. ºC.": 15,          # Temperature in Celsius
        "SED.": 16                # Sedimentation
    }
    
    # Process each data row (skip header row 0)
    for row_idx in range(1, len(df)):
        row = df.iloc[row_idx]
        
        # Get the description/product name from the first column
        descripcion = str(row[0]).strip() if pd.notna(row[0]) else ""
        
        # Skip empty rows or invalid descriptions
        if not descripcion or descripcion == "nan" or len(descripcion) < 2:
            continue
        
        # Skip separator rows (rows with only dashes, underscores, etc.)
        if all(c in '—-_=*\t ' for c in descripcion):
            continue
        
        # Create indicator with description
        indicator = {"DESCRIPCION": descripcion}
        
        # Add all metric values for this row
        for col_name, col_idx in col_mapping.items():
            # Check if column index is within bounds
            if col_idx < len(row):
                indicator[col_name] = parse_value(row[col_idx])
            else:
                indicator[col_name] = None
        
        indicators.append(indicator)
    
    return indicators


def parse_value(val):
    """
    Parse a value from Excel cell to appropriate Python type.
    
    This function converts Excel cell values to the most appropriate Python type:
    - Numbers are converted to int or float
    - Empty cells become None (will be null in JSON)
    - Text remains as string
    
    Args:
        val: Value from Excel cell (can be any type)
    
    Returns:
        Parsed value as int, float, string, or None
    """
    # Check if value is missing (NaN or None)
    if pd.isna(val):
        return None
    
    # Convert to string for processing
    val_str = str(val).strip()
    
    # Check if string is empty or "nan"
    if not val_str or val_str == "" or val_str.lower() == "nan":
        return None
    
    # Remove thousands separators (commas) for number parsing
    # Example: "1,234.56" becomes "1234.56"
    val_str = val_str.replace(',', '')
    
    # Try to parse as number
    try:
        # If string contains decimal point, parse as float
        if '.' in val_str:
            return float(val_str)
        else:
            # Otherwise, parse as integer
            return int(val_str)
    except ValueError:
        # If parsing as number fails, return as string
        # This handles text values like descriptions or labels
        return val_str if val_str else None


def main():
    """
    Main function to handle command-line execution.
    
    Usage:
        python excel-to-json.py <excel_file> [output_json]
    
    Examples:
        python excel-to-json.py 1003_page2_tables.xlsx
        python excel-to-json.py 1003_page2_tables.xlsx output.json
        python excel-to-json.py 1003_page2_tables.xlsx > output.json
    """
    # Check if required arguments are provided
    if len(sys.argv) < 2:
        # Print usage instructions if no arguments provided
        print("Usage:")
        print(f"    python {sys.argv[0]} <excel_file> [output_json]")
        print(f"")
        print(f"Arguments:")
        print(f"    excel_file   - Path to the Excel file to convert (required)")
        print(f"    output_json  - Output JSON file path (optional)")
        print(f"                   If not specified, uses Excel name with .json extension")
        print(f"")
        print(f"Examples:")
        print(f"    python {sys.argv[0]} 1003_page2_tables.xlsx")
        print(f"    python {sys.argv[0]} 1003_page2_tables.xlsx output.json")
        print(f"    python {sys.argv[0]} 1003_page2_tables.xlsx > output.json")
        print(f"")
        print(f"Note:")
        print(f"    Use pdf-to-excel.py first to extract tables from PDF to Excel")
        sys.exit(1)
    
    # Get command-line arguments
    excel_path = sys.argv[1]
    output_json = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Execute the conversion
    json_path = excel_to_json(excel_path, output_json)
    
    # Output the JSON content to stdout (allows piping to other commands or files)
    with open(json_path, 'r', encoding='utf-8') as f:
        print(f.read())


# Standard Python idiom to execute main() when script is run directly
if __name__ == "__main__":
    main()
