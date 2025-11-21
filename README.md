# CAMS Center Statement Extractor

A desktop application to extract and process transactions from CAMS Center Statement Excel files.

## Features

- User-friendly GUI built with ttkbootstrap
- Extracts transaction details including dates, clients, products, and pricing
- Automatically generates formatted Excel output files
- Cross-platform support (Windows, Linux, macOS)

## Installation

### Option 1: Run from Source

1. Install Python 3.11 or higher
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   python gui_settings.py
   ```

### Option 2: Download Executable

Download the pre-built Windows executable from the [Actions](../../actions) page:
- Go to the latest successful workflow run
- Download **CAMSExtractor-Windows** artifact
- Extract and run `CAMSExtractor.exe` - no Python installation required!

## Building Executable Locally

### Prerequisites
```bash
pip install -r requirements.txt
```

### Build
```bash
python build_executable.py
```

The executable will be created in the `dist/` folder.

### Manual PyInstaller Build
```bash
pyinstaller --noconfirm CAMSExtractor.spec
```

## GitHub Actions Auto-Build

This repository includes automated builds via GitHub Actions:

1. **Push to main/master**: Automatically builds Windows executable
2. **Manual trigger**: Go to Actions tab → Build Windows EXE → Run workflow

Build artifacts are available in the Actions tab for download.

## Usage

1. Launch the application
2. Click "Browse Excel" to select your CAMS Center Statement file
3. Click "Extract Transactions"
4. View progress in the logs
5. Find the output Excel file at the displayed path

## Input File Format

The application expects an Excel file with a sheet named `CamsCenterStatement` containing transaction data with the following structure:
- Date markers
- Client information
- Transaction times
- Product/service details
- Pricing and discount information
- Summary totals

## Output Format

The generated Excel file contains:
- Date
- Transaction Date & Time
- Client Name
- Reference ID
- Product/Service details
- Quantities, prices, discounts
- Subtotals, tax, and totals

## Development

### Project Structure
```
.
├── gui_settings.py                              # GUI application
├── extract_CamsCenterStatement_transactions.py  # Extraction logic
├── build_executable.py                          # Local build script
├── requirements.txt                             # Python dependencies
└── .github/workflows/build-executable.yml       # CI/CD pipeline
```

### Dependencies
- pandas
- openpyxl
- ttkbootstrap
- pyinstaller (for building executables)

## License

Copyright © 2025. All rights reserved.
