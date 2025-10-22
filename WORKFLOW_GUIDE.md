# Workflow Guide: PDF to JSON (No OCR)

## Overview

This guide describes the two-step process to extract tables from PDF without using OCR, providing an Excel intermediate step for human validation.

## Why No OCR?

- ❌ **OCR Issues**: OCR can misread numbers (e.g., 0 vs O, 1 vs l)
- ✅ **Direct Extraction**: pdfplumber extracts text directly from PDF structure
- ✅ **Human Validation**: Excel allows visual verification and corrections
- ✅ **Consistent Results**: Same PDF always produces same output

## Step-by-Step Workflow

### Step 1: Extract PDF to Excel

```bash
python extract_pdf_to_excel.py <pdf_file> [page_number]
```

**Example:**
```bash
python extract_pdf_to_excel.py 1003.pdf 2
```

**Output:**
- Creates `1003_page2_tables.xlsx` with 4 sheets:
  - `Raw_Data`: Complete raw extraction
  - `ANALISIS`: ANALISIS table data
  - `PRODUCTO_TERMINADO`: PRODUCTO TERMINADO table data
  - `CONTINUACION`: CONTINUACIÓN table data

### Step 2: Validate in Excel

1. Open the generated Excel file
2. Review each sheet for accuracy
3. Make any necessary corrections:
   - Fix misaligned values
   - Correct any obvious errors
   - Adjust number formats if needed
4. Save the file (keep the same name)

### Step 3: Convert Excel to JSON

```bash
python extract_pdf_to_excel.py --excel-to-json <excel_file> [output_json]
```

**Example:**
```bash
python extract_pdf_to_excel.py --excel-to-json 1003_page2_tables.xlsx output.json
```

**Output:**
- Creates JSON file with 4 groups:
  - `ANALISIS - HOY`
  - `ANALISIS - HASTA`
  - `PRODUCTO TERMINADO (AZUCAR)`
  - `PRODUCTO TERMINADO (AZUCAR) - CONTINUACIÓN`

## Complete Example

```bash
# Step 1: Extract to Excel
python extract_pdf_to_excel.py 1003.pdf 2
# Output: 1003_page2_tables.xlsx

# Step 2: Open and validate in Excel
# (Manual step - review and save)

# Step 3: Convert to JSON
python extract_pdf_to_excel.py --excel-to-json 1003_page2_tables.xlsx final.json
# Output: final.json
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
            },
            ...
        ]
    },
    {
        "Nombre Agrupador": "ANALISIS - HASTA",
        "Indicadores": [...]
    },
    {
        "Nombre Agrupador": "PRODUCTO TERMINADO (AZUCAR)",
        "Indicadores": [...]
    },
    {
        "Nombre Agrupador": "PRODUCTO TERMINADO (AZUCAR) - CONTINUACIÓN",
        "Indicadores": [...]
    }
]
```

## Advantages

1. **No OCR Errors**: Direct PDF text extraction
2. **Human Validation**: Excel provides familiar validation interface
3. **Easy Corrections**: Fix issues before JSON export
4. **Deterministic**: Same input = same output
5. **Faster**: No image preprocessing or OCR processing

## Troubleshooting

### Issue: Values in wrong columns

**Solution**: This is why we have the Excel step! Open the Excel file, review the data, and move values to correct columns before converting to JSON.

### Issue: Numbers appear as text

**Solution**: In Excel, select the cells and change format to Number before converting to JSON.

### Issue: Missing data

**Solution**: Check the Raw_Data sheet to see what was extracted. If data is missing from PDF structure, it cannot be extracted without OCR.

## Requirements

- Python 3.7+
- pandas
- openpyxl
- pdfplumber

Install with:
```bash
pip install -r requirements.txt
```

No OCR software required!
