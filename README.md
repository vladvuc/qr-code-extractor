# QR Code Sticker Detection

Python script to process Excel files containing image URLs, download images, detect QR codes, and write results to a new column.

## Prerequisites

- Python 3.8+
- Homebrew (macOS)

## Installation

1. Install system dependency (zbar):

```bash
brew install zbar
```

2. Create and activate virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate
```

3. Install Python dependencies:

```bash
pip install -r requirements.txt
```

## Usage

1. Ensure your Excel file `EXCEL/export_file.xlsx` is in the project directory

2. Activate the virtual environment (if not already active):

```bash
source venv/bin/activate
```

3. Run the `qr_processor` first:
   > [!tip]
   >
   > This will recognoze most of the QR codes, because it is using QR code reader, and it is the most recommended script, but it made few errors and it didn't classify everything

```bash
python qr_processor.py
```

Results will be saved to `output/company_export_processed_YYYYMMDD_HHMMSS.xlsx`

4. Run `process_qr_codes.py`second

> [!tip]
>
> The second script which is more using image analysis and qr code reader, it classified few more examples

Results will be saved to: `EXCEL/[name]_Analyzed.xsls`

5. Run `analyze_qr_codes.py` the last

> [!tip]
>
> This is the most basic script, but it managed to recognize the most of the left over QR codes.

Result will be saved to: `EXCEL/[name]\_Analyzed.xsls

## Output

The script will:

- Add a new column "QR CODE VALUE" to the Excel file
- For each image URL in the first column:
  - Download the image
  - Detect and decode any QR codes
  - Write the QR code value if found
  - Write "NOT A STICKER" if no QR code is detected or download fails

## Logs

Processing logs are saved to `logs/processing_YYYYMMDD_HHMMSS.log`

## Project Structure

```
.
├── qr_processor.py           # Main script
├── config.py                 # Configuration constants
├── requirements.txt          # Python dependencies
├── utils/
│   ├── __init__.py
│   ├── excel_handler.py     # Excel read/write operations
│   ├── image_downloader.py  # Image download with retry logic
│   └── qr_detector.py       # QR code detection/decoding
├── output/                   # Generated output files
└── logs/                     # Processing logs
```
