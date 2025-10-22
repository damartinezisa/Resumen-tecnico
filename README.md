# Table Extraction from Technical Reports

This project extracts tabular data from technical report PDFs **without using OCR**. Instead, it uses direct PDF table extraction for maximum accuracy.

## Approach

The extraction process follows a two-step workflow:

1. **PDF → Excel**: Extract tables directly from PDF and save to Excel for validation
2. **Excel → JSON**: Convert validated Excel data to structured JSON

This approach ensures:
- ✅ **No OCR errors** with numbers
- ✅ **Human validation** of extracted data via Excel
- ✅ **Maximum accuracy** through visual verification
- ✅ **Clean structured output** in JSON format

## Features

- **Direct PDF table extraction** using pdfplumber (no OCR)
- **Excel export** for easy data validation and correction
- **Formatted Excel output** with highlighted headers and auto-sized columns
- Extracts 3 main tables:
  1. ANALISIS (split into HOY and HASTA groups)
  2. PRODUCTO TERMINADO (AZUCAR)
  3. PRODUCTO TERMINADO (AZUCAR) - CONTINUACIÓN
- Outputs structured JSON data
- Handles empty cells (returns `null`)
- Pre-validation in Excel before final JSON export

## Requirements

- Python 3.7+

## Installation

Install Python dependencies:
```bash
pip install -r requirements.txt
```

No additional software needed - OCR tools are NOT required!

## Usage

### Step 1: Extract PDF to Excel

```bash
python extract_pdf_to_excel.py <pdf_file> [page_number]
```

This will create an Excel file with the extracted tables.

**Example:**
```bash
python extract_pdf_to_excel.py 1003.pdf 2
# Creates: 1003_page2_tables.xlsx
```

### Step 2: Validate in Excel

1. Open the generated Excel file (e.g., `1003_page2_tables.xlsx`)
2. Review the data in each sheet:
   - `Raw_Data`: Complete raw extraction
   - `ANALISIS`: ANALISIS table data
   - `PRODUCTO_TERMINADO`: PRODUCTO TERMINADO table data
   - `CONTINUACION`: CONTINUACIÓN table data
3. Make any necessary corrections
4. Save the file

### Step 3: Convert Excel to JSON

```bash
python extract_pdf_to_excel.py --excel-to-json <excel_file> [output_json]
```

This converts the validated Excel data to JSON format.

**Example:**
```bash
python extract_pdf_to_excel.py --excel-to-json 1003_page2_tables.xlsx output.json
# Creates: output.json
```

### Complete Workflow Example

```bash
# Step 1: Extract PDF to Excel
python extract_pdf_to_excel.py 1003.pdf 2

# Step 2: Validate data in Excel (open and review 1003_page2_tables.xlsx)

# Step 3: Convert to JSON
python extract_pdf_to_excel.py --excel-to-json 1003_page2_tables.xlsx > final_output.json
```

## Output Format

The JSON output follows this structure:

```json
[
    {
        "Nombre Agrupador": "ANALISIS - HOY",
        "Indicadores": [
            {
                "DESCRIPCION": "JUGO ABSOLUTO",
                "BRIX": 16.5,
                "SAC": 14.27,
                "PZA": 86.48,
                "COLOR": null,
                "PH": null,
                "GR": null
            }
        ]
    }
]
```

## Advantages Over OCR

- **No number recognition errors**: PDF text is extracted directly, not read from images
- **Faster processing**: No image preprocessing or OCR engine needed
- **Human validation**: Excel provides a familiar interface for data verification
- **Easy corrections**: Fix any issues directly in Excel before JSON export
- **Consistent results**: Same input always produces same output (deterministic)

## How It Works

### Step 1: PDF Table Extraction

1. **Direct Table Detection**: Uses pdfplumber to detect table boundaries directly from PDF structure
2. **No OCR Required**: Extracts text and numbers directly from PDF (not from images)
3. **Table Separation**: Identifies and separates the 3 logical tables based on keywords
4. **Excel Export**: Saves each table to a separate sheet with formatting for easy review

### Step 2: Human Validation

1. **Visual Review**: Open Excel file and review extracted data
2. **Corrections**: Make any necessary corrections to numbers or text
3. **Quality Assurance**: Ensures 100% accuracy before JSON conversion

### Step 3: JSON Conversion

1. **Schema Mapping**: Converts Excel data to the required JSON structure
2. **Type Handling**: Automatically converts numbers to int/float, preserves null values
3. **Structured Output**: Formats data according to specification with proper groupings

## Notes

- The script extracts tables directly from PDF structure (not OCR-based)
- Excel files provide an intermediate validation step for accuracy
- Make corrections in Excel before converting to JSON
- Empty cells are represented as `null` in the JSON output
- The extraction is deterministic - same PDF always produces same Excel output
- Works best with PDFs that have structured tables (not scanned images)

## Development

The main components are:

- `extract_pdf_to_excel()`: Extracts tables from PDF page and saves to Excel
- `excel_to_json()`: Converts validated Excel file to JSON format
- `parse_analisis_from_df()`: Parses ANALISIS table from DataFrame
- `parse_producto_from_df()`: Parses PRODUCTO TERMINADO table from DataFrame
- `parse_continuacion_from_df()`: Parses CONTINUACIÓN table from DataFrame
- `parse_value()`: Converts Excel cell values to appropriate Python types (int/float/None)

## Troubleshooting

**Issue**: No tables found in PDF
- **Solution**: Ensure the PDF has actual table structures (not just images). The PDF should contain selectable text.

**Issue**: Data in wrong columns in Excel
- **Solution**: This is expected for complex tables. The Excel intermediate step allows you to manually correct the column alignment before JSON conversion.

**Issue**: Numbers appear as text in Excel
- **Solution**: The script tries to auto-detect numbers. You can manually correct cell types in Excel before converting to JSON.

## License

This project is provided as-is for educational and research purposes.
