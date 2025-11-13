# Army Height and Weight PDF Generator

Automates the generation of Army height and weight assessment forms (DA 5500/5501) from Excel data, with support for AR 600-9 compliance tracking and AFT exemptions.

## Process Overview

### Step 1: Setup
1. **Download the complete package** containing:
   - `generate_pdfs.py` (main script)
   - DA 5500 PDF template (male soldiers)
   - DA 5501 PDF template (female soldiers)
   - Excel template with built-in calculations

### Step 2: Conduct Height and Weight Assessment
1. **Open the Excel template** and fill out soldier information
2. **Enter basic data**: Name, rank, gender, age, height, weight
3. **Excel automatically calculates**:
   - Maximum allowable weight based on age/height lookup tables
   - Pass/Fail status for height and weight standards
   - Immediate feedback on whether soldier needs tape test

### Step 3: Tape Test (if required)
1. **If soldier needs taping**: Excel will indicate "Needs Tape"
2. **Conduct tape measurements** and enter the 3 measurements from NCOs/Officers
3. **Excel automatically calculates**:
   - Average of the 3 measurements
   - Body fat percentage using Army regulation formulas
   - Maximum allowable body fat percentage from lookup tables
   - Pass/Fail status for body fat standards

### Step 4: Personnel Information
1. **Set default personnel** at top of Excel:
   - Row 2: Default "Prepared By" name, rank, initials
   - Row 4: Default "Supervised By" name, rank
2. **For soldiers requiring tape test**: Override defaults by entering specific personnel in columns R-U for that soldier's row

### Step 5: Generate PDFs
1. **Save the Excel file** after completing all entries
2. **Run the script**:
   ```bash
   python generate_pdfs.py soldiers.xlsx DA5500_template.pdf DA5501_template.pdf ./output
   ```
3. **Script automatically**:
   - Creates individual PDFs (5500 for males, 5501 for females)
   - Fills forms with soldier data and measurements
   - Generates appropriate AR 600-9 compliance remarks
   - Handles AFT exemptions for qualifying soldiers

### Step 6: Final Processing
1. **Print generated PDFs**
2. **Obtain required signatures**
3. **File completed forms** per unit procedures

## Future Enhancements

### Planned Features
- **DTMS Integration**: Export data in format for bulk DTMS updates
- **Name Matching**: Automated soldier name formatting to match DTMS requirements

### Known Challenges
- Soldier names must match DTMS database exactly for bulk updates
- Name formatting inconsistencies may require manual verification

## Technical Requirements

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
| G | AFT Pass | "Yes" for AFT exemption |
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
- **AFT exemptions**: Handles soldiers who pass AFT but fail tape test
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
- AFT exemption criteria (540+ total score, 80+ per event)