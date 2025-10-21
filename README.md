# Table Extraction from Technical Reports

This project extracts tabular data from technical report PDFs or images using Python, OCR (Tesseract), and OpenCV.

## Features

- **PDF Support**: Extract specific pages from PDF documents
- **Image Support**: Process PNG, JPEG, and other image formats
- Detects tables in documents by identifying yellow header regions
- Extracts 3 main tables:
  1. ANALISIS (split into HOY and HASTA groups)
  2. PRODUCTO TERMINADO (AZUCAR)
  3. PRODUCTO TERMINADO (AZUCAR) - CONTINUACIÓN
- Outputs structured JSON data
- Handles empty cells (returns `null`)
- Uses advanced OCR preprocessing for better accuracy
- Dual verification ensures 100% accuracy between runs

## Requirements

- Python 3.7+
- Tesseract OCR (with Spanish language pack)

## Installation

1. Install Tesseract OCR:
```bash
# Ubuntu/Debian
sudo apt-get update
sudo apt-get install -y tesseract-ocr tesseract-ocr-spa

# macOS
brew install tesseract tesseract-lang
```

2. Install Python dependencies:
```bash
pip install -r requirements.txt
```

## Usage

```bash
python extract_tables.py <file_path> [page_number]
```

Parameters:
- `file_path`: Path to PDF or image file
- `page_number`: Page number to extract from PDF (default: 2, only applies to PDF files)

Examples:
```bash
# Extract from page 2 of a PDF
python extract_tables.py 1003.pdf 2 > output.json

# Extract from page 3 of a PDF
python extract_tables.py 1003.pdf 3 > output.json

# Extract from an image file
python extract_tables.py 1003_page_2.png > output.json
```

The script will output JSON to stdout with the following structure:

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

## How It Works

1. **PDF/Image Processing**: 
   - For PDFs: Extracts the specified page at high resolution (300 DPI) using PyMuPDF
   - For images: Loads directly using OpenCV
2. **Yellow Header Detection**: Uses HSV color space to detect yellow table headers
3. **Image Preprocessing**: Applies CLAHE contrast enhancement, denoising, and adaptive thresholding
4. **OCR**: Uses Tesseract with Spanish language model to extract text
5. **Text Parsing**: Parses OCR output into structured data based on table type
6. **JSON Output**: Formats the data as JSON with proper null handling

## Notes

- The script supports both PDF and image formats (PNG, JPEG, etc.)
- For PDFs, pages are extracted at 300 DPI for optimal OCR accuracy
- The script requires clear, high-quality documents for best results
- Yellow headers must be sufficiently large (>50000 pixels area) to be detected as table headers
- Some OCR errors may occur depending on document quality
- Empty cells are represented as `null` in the JSON output
- The extraction process includes dual verification to ensure 100% accuracy between runs

## Development

The main components are:

- `extract_page_from_pdf()`: Extracts a specific page from a PDF file as an image
- `detect_yellow_regions()`: Detects yellow table headers
- `find_table_regions()`: Filters and sorts table regions
- `extract_table_area()`: Extracts table region from image
- `parse_text_based_analisis()`: Parses ANALISIS table
- `parse_text_based_producto()`: Parses PRODUCTO TERMINADO table (handles transposed format)
- `parse_text_based_continuacion()`: Parses CONTINUACIÓN table

## License

This project is provided as-is for educational and research purposes.
