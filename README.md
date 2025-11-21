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

Download the pre-built executable for your platform from the [Releases](../../releases) page:
- **Windows**: `CAMSExtractor-Windows.zip`
- **Linux**: `CAMSExtractor-Linux.zip`
- **macOS**: `CAMSExtractor-macOS.zip`

Extract and run the executable directly - no Python installation required!

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
# Windows
pyinstaller --onefile --windowed --name "CAMSExtractor" gui_settings.py

# Linux/macOS
pyinstaller --onefile --windowed --name "CAMSExtractor" gui_settings.py
```

## GitHub Actions Auto-Build

This repository includes automated builds via GitHub Actions:

1. **Push to main/master**: Builds executables for all platforms
2. **Create a tag**: Builds and creates a GitHub Release
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
3. **Manual trigger**: Go to Actions tab → Build Executable → Run workflow

Artifacts are available in the Actions tab for 30 days, or permanently in Releases.

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
