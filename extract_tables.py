#!/usr/bin/env python3
"""
Script to extract tables from technical report images.
Extracts 3 tables with yellow headers and outputs structured JSON data.
"""

import cv2
import numpy as np
import pytesseract
from PIL import Image
import json
import re
import sys
from typing import List, Dict, Any

def detect_yellow_regions(image):
    """Detect yellow regions in the image (table headers)"""
    # Convert to HSV for better color detection
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    
    # Define range for yellow color
    # Yellow in HSV: H: 20-40, S: 100-255, V: 100-255
    lower_yellow = np.array([20, 80, 80])
    upper_yellow = np.array([40, 255, 255])
    
    # Create mask for yellow regions
    mask = cv2.inRange(hsv, lower_yellow, upper_yellow)
    
    return mask

def find_table_regions(image, yellow_mask):
    """Find table regions based on yellow headers"""
    # Find contours of yellow regions
    contours, _ = cv2.findContours(yellow_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # Filter and sort contours by y-coordinate (top to bottom)
    table_regions = []
    for contour in contours:
        area = cv2.contourArea(contour)
        # Filter for large yellow regions (table headers), not small icons/labels
        if area > 50000:
            x, y, w, h = cv2.boundingRect(contour)
            table_regions.append({'x': x, 'y': y, 'w': w, 'h': h, 'area': area})
    
    # Sort by y coordinate (top to bottom)
    table_regions = sorted(table_regions, key=lambda r: r['y'])
    
    return table_regions

def extract_table_area(image, header_region, expansion=200):
    """Extract table area including rows below the header"""
    height, width = image.shape[:2]
    
    # Use the full header width plus some padding
    x = max(0, header_region['x'] - 20)
    y = header_region['y'] - 5  # Include a bit above header
    
    # Ensure we capture the full table width
    x_end = header_region['x'] + header_region['w'] + 20
    w = min(width, x_end) - x
    
    h = min(height - y, expansion + 50)
    
    return image[y:y+h, x:x+w]

def clean_text(text):
    """Clean OCR text output"""
    if not text:
        return None
    # Remove extra whitespace
    text = re.sub(r'\s+', ' ', text.strip())
    # Remove special characters that are OCR artifacts
    text = text.replace('|', '').replace('_', '')
    return text if text else None

def parse_number(text):
    """Parse numeric values from text"""
    if not text or text == '':
        return None
    
    text = str(text).strip()
    if not text or text.lower() in ['null', 'none', '', '-', 'n/a']:
        return None
    
    # Remove commas and clean
    text = text.replace(',', '').replace(' ', '')
    
    try:
        # Try to parse as float
        if '.' in text:
            return float(text)
        else:
            return int(text)
    except ValueError:
        # If parsing fails, check if it's a quoted string
        if text.startswith('"') and text.endswith('"'):
            return text
        return None





def parse_text_based_analisis(text):
    """Parse ANALISIS table from raw text"""
    hoy_indicators = []
    hasta_indicators = []
    
    lines = text.strip().split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Remove pipe characters (table borders detected by OCR)
        line = line.replace('|', ' ')
        
        # Skip header/title lines (but not data rows with those words)
        upper_line = line.upper()
        if 'HOY' in upper_line and 'HASTA' in upper_line and len(line.split()) < 10:
            continue
        if any(skip in upper_line for skip in ['ANALISIS', 'BRIX', 'SAC', 'PZA', 'COLOR', 'PH']):
            if not any(word in upper_line for word in ['JUGO', 'CACHAZA', 'MELADURA', 'SIROPE', 'MAGMA', 'RUN', 'MASA', 'MIEL']):
                continue
        
        # Split into tokens
        tokens = line.split()
        if len(tokens) < 2:
            continue
        
        # Find where numeric values start
        desc_parts = []
        values = []
        found_number = False
        
        for token in tokens:
            # Check if token is numeric (including decimals and commas)
            if re.match(r'^[\d,\.]+$', token):
                found_number = True
                # Clean and add to values
                values.append(token.replace(',', ''))
            else:
                if not found_number:
                    desc_parts.append(token)
                # If we find text after numbers, it might be part of description
                # but likely noise, so ignore
        
        if not desc_parts:
            continue
        
        description = ' '.join(desc_parts)
        
        # Create indicators
        hoy_indicator = {"DESCRIPCION": description}
        hasta_indicator = {"DESCRIPCION": description}
        
        # HOY columns: BRIX, SAC, PZA (first 3 values)
        # HASTA columns: BRIX, SAC, PZA (next 3 values)
        # Then HOY: COLOR, PH, GR (next 3 values)
        
        hoy_cols = ["BRIX", "SAC", "PZA", "COLOR", "PH", "GR"]
        hasta_cols = ["BRIX", "SAC", "PZA"]
        
        # Parse first 3 for HOY BRIX, SAC, PZA
        for i in range(3):
            if i < len(values):
                hoy_indicator[hoy_cols[i]] = parse_number(values[i])
            else:
                hoy_indicator[hoy_cols[i]] = None
        
        # Parse next 3 for HASTA BRIX, SAC, PZA
        for i in range(3):
            idx = 3 + i
            if idx < len(values):
                hasta_indicator[hasta_cols[i]] = parse_number(values[idx])
            else:
                hasta_indicator[hasta_cols[i]] = None
        
        # Parse remaining for HOY COLOR, PH, GR
        for i in range(3, 6):
            idx = 6 + (i - 3)
            col = hoy_cols[i]
            if idx < len(values):
                hoy_indicator[col] = parse_number(values[idx])
            else:
                hoy_indicator[col] = None
        
        hoy_indicators.append(hoy_indicator)
        hasta_indicators.append(hasta_indicator)
    
    return hoy_indicators, hasta_indicators

def parse_text_based_producto(text):
    """Parse PRODUCTO TERMINADO table from raw text"""
    indicators = []
    
    lines = text.strip().split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        line = line.replace('|', ' ')
        
        # Stop at CONTINUACIÓN section
        if 'DESCRIPCION' in line.upper() and 'Quintales' in line:
            break
        
        # Skip header/title lines
        if any(skip in line.upper() for skip in ['PRODUCTO', 'TERMINADO', 'AZUCAR']):
            if 'QQ' not in line.upper() and 'AZ.' not in line.upper() and 'TOTAL' not in line.upper():
                continue
        
        # Skip column header line
        if line.upper().startswith('ESTANDAR') and 'DIA' in line.upper() and 'HOY' in line.upper():
            continue
        
        # Split into tokens
        tokens = line.split()
        if len(tokens) < 2:
            continue
        
        # Find where numeric values start
        desc_parts = []
        values = []
        found_number = False
        
        for token in tokens:
            # Check if token is numeric
            if re.match(r'^[\d,\.]+$', token):
                found_number = True
                values.append(token.replace(',', ''))
            else:
                if not found_number:
                    desc_parts.append(token)
        
        if not desc_parts:
            continue
        
        description = ' '.join(desc_parts)
        
        # Skip if description looks like a column header
        if description.upper() in ['ESTANDAR/DIA', 'ESTANDARDIA', 'HOY', 'HASTA']:
            continue
        
        indicator = {"DESCRIPCION": description}
        
        # Columns: ESTANDAR/DIA, HOY, HASTA (3 values)
        columns = ["ESTANDAR/DIA", "HOY", "HASTA"]
        for i, col in enumerate(columns):
            if i < len(values):
                indicator[col] = parse_number(values[i])
            else:
                indicator[col] = None
        
        indicators.append(indicator)
    
    return indicators

def parse_text_based_continuacion(text):
    """Parse CONTINUACIÓN table from raw text"""
    indicators = []
    
    lines = text.strip().split('\n')
    
    column_names = ["DESCRIPCION", "Quintales", "% HUM.", "% CEN.", "% POL.", "COLOR", 
                    "Mg/kg Vit.\"A\"", "T. GRANO", "% C.V", "FS", "TEMP. ºC.", "SED."]
    
    # Find where the CONTINUACIÓN section starts
    start_parsing = False
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        line = line.replace('|', ' ')
        
        # Check if we've reached the CONTINUACIÓN section header
        if 'DESCRIPCION' in line and 'Quintales' in line:
            start_parsing = True
            continue
        
        if not start_parsing:
            continue
        
        # Skip column header lines
        if any(skip in line.upper() for skip in ['CONTINUACION', 'CONTINUACIÓN', 'DESCRIPCION']):
            if 'Quintales' in line or 'HUM' in line:
                continue
        
        # Split into tokens
        tokens = line.split()
        if len(tokens) < 2:
            continue
        
        # Find where numeric values start
        desc_parts = []
        values = []
        found_number = False
        
        for token in tokens:
            # Check if token is numeric (including negative signs for decimals represented with -)
            if re.match(r'^-?[\d,\.\"]+$', token):
                found_number = True
                cleaned = token.replace(',', '').replace('-', '.')
                values.append(cleaned)
            else:
                if not found_number:
                    desc_parts.append(token)
        
        if not desc_parts:
            continue
        
        description = ' '.join(desc_parts)
        
        # Skip lines that are separators
        if all(c in '—-_= \t' for c in description):
            continue
        
        indicator = {"DESCRIPCION": description}
        
        # Parse 11 columns after description
        for i, col in enumerate(column_names[1:]):
            if i < len(values):
                value = parse_number(values[i])
                # Special handling for FS column
                if col == "FS" and value is None and i < len(values):
                    raw_val = values[i].strip('"')
                    value = raw_val if raw_val else None
                indicator[col] = value
            else:
                indicator[col] = None
        
        indicators.append(indicator)
    
    return indicators

def main(image_path):
    """Main function to extract tables from image"""
    # Read image
    image = cv2.imread(image_path)
    if image is None:
        print(f"Error: Could not read image {image_path}", file=sys.stderr)
        return None
    
    # Detect yellow regions
    yellow_mask = detect_yellow_regions(image)
    
    # Find table regions
    table_regions = find_table_regions(image, yellow_mask)
    
    if len(table_regions) < 3:
        print(f"Warning: Expected 3 tables but found {len(table_regions)}", file=sys.stderr)
    
    results = []
    
    # Process each table
    for idx, region in enumerate(table_regions):
        # Extract table area (expand downward to capture all rows)
        if idx == 0:  # First table (ANALISIS)
            table_img = extract_table_area(image, region, expansion=450)
        elif idx == 1:  # Second table (PRODUCTO TERMINADO)
            table_img = extract_table_area(image, region, expansion=350)
        else:  # Third table (CONTINUACIÓN)
            table_img = extract_table_area(image, region, expansion=450)
        
        # Perform OCR
        gray = cv2.cvtColor(table_img, cv2.COLOR_BGR2GRAY)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        gray = clahe.apply(gray)
        gray = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
        thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                         cv2.THRESH_BINARY, 11, 2)
        
        config = '--psm 6'
        text = pytesseract.image_to_string(thresh, lang='spa', config=config)
        
        # Parse based on table type using text-based parsing
        if idx == 0:  # ANALISIS table
            hoy_indicators, hasta_indicators = parse_text_based_analisis(text)
            results.append({
                "Nombre Agrupador": "ANALISIS - HOY",
                "Indicadores": hoy_indicators
            })
            results.append({
                "Nombre Agrupador": "ANALISIS - HASTA",
                "Indicadores": hasta_indicators
            })
        
        elif idx == 1:  # PRODUCTO TERMINADO (AZUCAR)
            indicators = parse_text_based_producto(text)
            results.append({
                "Nombre Agrupador": "PRODUCTO TERMINADO (AZUCAR)",
                "Indicadores": indicators
            })
        
        elif idx == 2:  # CONTINUACIÓN
            indicators = parse_text_based_continuacion(text)
            results.append({
                "Nombre Agrupador": "PRODUCTO TERMINADO (AZUCAR) - CONTINUACIÓN",
                "Indicadores": indicators
            })
    
    return results

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_tables.py <image_path>")
        sys.exit(1)
    
    image_path = sys.argv[1]
    results = main(image_path)
    
    if results:
        # Output JSON without extra text
        print(json.dumps(results, indent=4, ensure_ascii=False))
    else:
        print("Error: Could not extract tables", file=sys.stderr)
        sys.exit(1)
