# Ajman University - Form Filler

A Python GUI application for generating PDF forms for Ajman University Main Store.

## Features

- **Two Form Types:**
  - Acknowledgment of Receipt Form
  - Asset Transfer Form (ATF)

- **Dynamic Rows:** Add/remove items or assets as needed
- **Excel Import:** Batch import items from Excel files (.xlsx)
- **Auto-fill Date:** Current date automatically populated
- **Digital Signature:** Adobe Acrobat compatible certificate-based signature fields
- **Smart Filenames:** Auto-generated filenames with duplicate handling (#2, #3, etc.)

## Screenshots

The application provides a tabbed interface for easy form selection:
- Acknowledgment of Receipt tab
- Asset Transfer Form tab

## Installation

### Requirements
- Python 3.8+
- Dependencies listed in `requirements.txt`

### Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/form-filler.git
   cd form-filler
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python src/main.py
   ```

   Or on Windows, double-click `run.bat`

## Usage

### Manual Entry
1. Select the form type (Acknowledgment or Transfer)
2. Fill in the required fields
3. Add items/assets using the "+ Add Item" or "+ Add Asset" button
4. Click "Generate PDF"

### Excel Import
1. Click "Import from Excel" button
2. Select your Excel file
3. Items will be automatically populated

Sample Excel files are provided in the `samples/` folder.

### Excel File Format

**Acknowledgment Form:**
| Store Code | Item Description | Qty | Purchase Date/LPO |
|------------|------------------|-----|-------------------|
| AU-IT-001  | Dell Desktop     | 1   | 15/01/2024        |

**Transfer Form:**
| Store Code | Asset Name | Description | Old Asset No. |
|------------|------------|-------------|---------------|
| AU-IT-001  | Dell Desktop | Intel i7, 16GB | OLD-12345 |

## Output Files

Generated PDFs are saved in the `output/` folder with automatic naming:

- **Acknowledgment:** `{Emp ID} - {Name} - acknowledgement form {Asset}.pdf`
- **Transfer:** `Asset Transfer - From {ID}-{Name} to {ID}-{Name}.pdf`

## Building Standalone Executable

To create a standalone `.exe` that doesn't require Python:

1. On Windows, run:
   ```bash
   build.bat
   ```

2. The executable will be created in `dist/FormFiller.exe`

## Project Structure

```
form-filler/
├── src/
│   ├── main.py              # Application entry point
│   ├── gui/
│   │   ├── main_window.py   # Main tabbed interface
│   │   ├── acknowledgment_form.py
│   │   └── transfer_form.py
│   ├── pdf/
│   │   ├── acknowledgment_pdf.py
│   │   └── transfer_pdf.py
│   └── utils/
│       └── signature.py
├── Forms/                    # Logo and reference files
├── samples/                  # Sample Excel files
├── output/                   # Generated PDFs
├── requirements.txt
├── build.py                  # PyInstaller build script
├── run.bat                   # Windows launcher
└── README.md
```

## Dependencies

- `reportlab` - PDF generation
- `pypdf` - PDF manipulation and signature fields
- `Pillow` - Image handling
- `openpyxl` - Excel file support
- `tkinter` - GUI (included with Python)

## License

MIT License

## Author

Developed for Ajman University Main Store
