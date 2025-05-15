# PDF Table OCR Extractor

A Python tool that automatically extracts tables from PDF files using OCR (Optical Character Recognition). It processes all PDFs in an input directory and outputs the extracted data in both CSV and Excel formats.

## Features

- Batch processing of multiple PDF files
- Automatic OCR using Tesseract
- Supports PDF rotation and cropping
- Customizable column selection
- Outputs both CSV and Excel files
- Progress bar for tracking processing status
- Handles German and English text

## Requirements

- Python 3.8 or higher
- Tesseract OCR
- Required Python packages (see `requirements.txt`)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/pdf-table-ocr-extractor.git
cd pdf-table-ocr-extractor
```

2. Create and activate a virtual environment:
```bash
python -m venv .venv
# On Windows:
.venv\Scripts\activate
# On Unix or MacOS:
source .venv/bin/activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

4. Install Tesseract OCR:
   - Windows: Download and install from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
   - Linux: `sudo apt-get install tesseract-ocr`
   - MacOS: `brew install tesseract`

## Usage

1. Create an `input` directory and place your PDF files there
2. Run the script:
```bash
python process_pdf_tables.py [--rotate ROTATE] [--crop CROP] [--columns COLUMNS]
```

### Parameters

- `--rotate`: Rotation angle in degrees (0, 90, 180, 270). Default: 0
- `--crop`: Right margin crop ratio (0-1). Default: 0.05
- `--columns`: Comma-separated list of columns to extract. Default: "Company,City,StreetNo,PostalCode,Name,Title,Email,Phone"

### Available Columns

The following columns are available by default:
- `Company`: Company name
- `City`: City name
- `StreetNo`: Street and number
- `PostalCode`: Postal code
- `Name`: Person's name
- `Title`: Person's title
- `Email`: Email address
- `Phone`: Phone number

You can specify any subset of these columns or add your own custom columns.

### Examples

Extract all default columns:
```bash
python process_pdf_tables.py
```

Extract only company and email:
```bash
python process_pdf_tables.py --columns "Company,Email"
```

Extract default columns with custom rotation:
```bash
python process_pdf_tables.py --rotate 90
```

### Output

The script creates an `output` directory containing:
- CSV files with all extracted data including page numbers
- Excel files with the same data (excluding page numbers)

## Project Structure

```
pdf-table-ocr-extractor/
├── input/                  # Place PDF files here
├── output/                 # Generated output files
├── .venv/                  # Virtual environment
├── process_pdf_tables.py   # Main script
├── requirements.txt        # Python dependencies
└── README.md              # This file
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details. 