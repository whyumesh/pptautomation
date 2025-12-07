# AIL LT PowerPoint Automation

Automatically generate monthly AIL LT PowerPoint presentations from Excel files using a template-based approach. This tool preserves exact formatting, positioning, and theme from your manual template while updating data from Excel files.

## Features

- ✅ **100% Formatting Match** - Uses template PPT as base, preserving all formatting
- ✅ **Exact Positioning** - Every element in the same position as template
- ✅ **Theme Preservation** - Complete theme and styling maintained
- ✅ **Charts/Images Preserved** - All visual elements kept intact
- ✅ **Monthly Automation** - Generate presentations for any month
- ✅ **Data Accuracy** - Extracts and populates data from Excel files

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Prepare Your Files

- **Template PPT**: `AIL LT - Sep'25.pptx` (must be in root directory)
- **Excel Files**: Place in `excel_files/` folder:
  - `AIL LT Working file.xlsx` (contains CLT and consent sheets)
  - `Chronic Missing Report AIL - [Month Range].xlsx` (contains missed HCP data)
  - `Overlapped Vacant deactivation - greater than 2.xlsb` (contains overlap data)

### 3. Generate Presentation

```bash
# For a specific month
python ail_lt_template_replicator.py --month "Oct'25"

# For current month
python ail_lt_template_replicator.py
```

## Usage

### Basic Usage

```bash
python ail_lt_template_replicator.py --month "Nov'25"
```

### With Custom Output Name

```bash
python ail_lt_template_replicator.py --month "Nov'25" --output-name "AIL LT - Nov'25_FINAL.pptx"
```

### With Custom Template

```bash
python ail_lt_template_replicator.py --month "Nov'25" --template "path/to/your/template.pptx"
```

### Command Line Options

```
--input-dir      Directory containing Excel files (default: excel_files)
--output-dir     Output directory for PPT file (default: output)
--month          Month (e.g., "Sep'25", "October 2025")
--template       Path to template PPT file (default: AIL LT - Sep'25.pptx)
--output-name    Custom output filename (optional)
```

## Project Structure

```
excel-to-ppt-automation/
├── ail_lt_template_replicator.py  # Main automation script
├── requirements.txt               # Python dependencies
├── README.md                      # This file
├── TEMPLATE_REPLICATOR_GUIDE.md  # Detailed usage guide
├── AIL LT - Sep'25.pptx          # Template PPT file
├── excel_files/                   # Excel data files
│   ├── AIL LT Working file.xlsx
│   ├── Chronic Missing Report AIL - [Range].xlsx
│   └── Overlapped Vacant deactivation - greater than 2.xlsb
└── output/                        # Generated PPT files
```

## How It Works

1. **Loads Template** - Uses your manual template PPT as the base
2. **Preserves Everything** - Keeps all formatting, positioning, images, charts
3. **Updates Data** - Populates tables from Excel files
4. **Updates Month/Year** - Changes title slide to new month
5. **Saves Output** - Creates new PPT with updated data

## Output

The generated PowerPoint will be saved as:
- **Filename:** `AIL LT - [Month]'[Year].pptx`
- **Location:** `output/` folder

Example: `output/AIL LT - Oct'25.pptx`

## What Gets Updated

- ✅ **Slide 1**: Month/year in title (e.g., "Oct|25")
- ✅ **Slide 3**: Project FMV table data
- ✅ **Slide 4**: Consent Status table data
- ✅ **Slide 6**: HCP Overlap table data
- ✅ **Slide 7**: Missed HCP table data

## What Gets Preserved

- ✅ **All formatting** (fonts, colors, sizes)
- ✅ **All positioning** (exact element positions)
- ✅ **All images and charts** (from template)
- ✅ **All theme elements** (backgrounds, styles)
- ✅ **All text content** (except data tables)

## Requirements

- Python 3.7+
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- python-pptx >= 0.6.21
- pyxlsb (for .xlsb file support)

## Monthly Workflow

1. Update Excel files for the new month in `excel_files/` folder
2. Run the script:
   ```bash
   python ail_lt_template_replicator.py --month "Nov'25"
   ```
3. Review the output - it will match the template exactly
4. Update charts if needed (charts from template are preserved)

## Troubleshooting

### File Not Found Error

If you get "Chronic Missing Report AIL - Jun to Aug.xlsx not found":
- The script looks for files with month ranges in the filename
- Rename your file to match the pattern, or update the script to search by pattern

### Template Not Found

Ensure `AIL LT - Sep'25.pptx` is in the root directory, or use `--template` to specify the path.

### Data Not Updating

- Check that Excel files are in `excel_files/` folder
- Verify sheet names match: 'CLT', 'consent', 'New Visual'
- Check that column names match expected names

## Documentation

- **TEMPLATE_REPLICATOR_GUIDE.md** - Detailed usage guide with examples

## Status

✅ **Production Ready** - The script is ready for monthly use and produces presentations that are indistinguishable from manually created ones (except for the updated data).
