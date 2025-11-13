# Army Height and Weight PDF Generator

Automates the generation of Army height and weight assessment forms (DA 5500/5501) from Excel data, with support for AR 600-9 compliance tracking and ACFT exemptions.

## Requirements

- Python 3.6+
- Required packages:
  ```bash
  pip install openpyxl fillpdf
  ```

## Usage

```bash
python generate_pdfs.py <excel_file> <pdf_5500_template> <pdf_5501_template> <output_directory> [options]
```

### Arguments

- `excel_file`: Excel file containing soldier data
- `pdf_5500_template`: PDF template for male soldiers (DA 5500)
- `pdf_5501_template`: PDF template for female soldiers (DA 5501)
- `output_directory`: Directory where generated PDFs will be saved

### Options

- `--date YYYYMMDD`: Custom date for forms (default: today)
- `--debug`: Enable debug output

## Excel File Format

The Excel file should contain soldier data starting from row 6 with the following columns:

| Column | Field | Description |
|--------|-------|-------------|
| A | Name | Soldier's name |
| B | Rank | Military rank |
| C | Gender | M/F (determines which template to use) |
| D | Age | Soldier's age |
| E | Height | Height measurement |
| F | Weight | Weight measurement |
| G | ACFT Pass | "Yes" for ACFT exemption |
| I | Height/Weight Status | "Pass" or "Needs Tape" |
| J-L | Tape Measurements | First, Second, Third measurements |
| M | Average | Average of tape measurements |
| N | Body Fat % | Calculated body fat percentage |
| Q | Fail Tape | "Fail Tape" status |
| R-U | Optional overrides | Prepared by, rank, approved by, rank |

### Default Values (Rows 2 and 4)

- Row 2: Default preparer name, rank, initials
- Row 4: Default supervisor name, rank

## Features

- **Automatic form selection**: Uses DA 5500 for males, DA 5501 for females
- **AR 600-9 compliance**: Generates appropriate remarks based on standards
- **ACFT exemptions**: Handles soldiers who pass ACFT but fail tape test
- **Error handling**: Validates files and handles missing data gracefully
- **Batch processing**: Processes multiple soldiers from single Excel file

## Output

Generated PDFs are named: `{SoldierName}_5500.pdf` or `{SoldierName}_5501.pdf`

## Example

```bash
python generate_pdfs.py soldiers.xlsx DA5500_template.pdf DA5501_template.pdf ./output --date 20241201
```

## Remarks Generated

The script automatically generates appropriate remarks based on:
- Height/weight pass/fail status
- Body fat percentage compliance
- ACFT exemption criteria (540+ total score, 80+ per event)