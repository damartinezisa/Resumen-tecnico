# Table Extraction from Technical Reports

This project extracts tabular data from technical report images using Python, OCR (Tesseract), and OpenCV.

## Features

- Detects tables in images by identifying yellow header regions
- Extracts 3 main tables:
  1. ANALISIS (split into HOY and HASTA groups)
  2. PRODUCTO TERMINADO (AZUCAR)
  3. PRODUCTO TERMINADO (AZUCAR) - CONTINUACIÓN
- Outputs structured JSON data
- Handles empty cells (returns `null`)
- Uses advanced OCR preprocessing for better accuracy

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
python extract_tables.py <image_path>
```

Example:
```bash
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

1. **Yellow Header Detection**: The script uses HSV color space to detect yellow table headers
2. **Image Preprocessing**: Applies CLAHE contrast enhancement, denoising, and adaptive thresholding
3. **OCR**: Uses Tesseract with Spanish language model to extract text
4. **Text Parsing**: Parses OCR output into structured data based on table type
5. **JSON Output**: Formats the data as JSON with proper null handling

## Notes

- The script requires clear, high-quality images for best results
- Yellow headers must be sufficiently large (>50000 pixels area) to be detected as table headers
- Some OCR errors may occur depending on image quality
- Empty cells are represented as `null` in the JSON output

## Development

The main components are:

- `detect_yellow_regions()`: Detects yellow table headers
- `find_table_regions()`: Filters and sorts table regions
- `extract_table_area()`: Extracts table region from image
- `parse_text_based_analisis()`: Parses ANALISIS table
- `parse_text_based_producto()`: Parses PRODUCTO TERMINADO table
- `parse_text_based_continuacion()`: Parses CONTINUACIÓN table

## License

This project is provided as-is for educational and research purposes.
