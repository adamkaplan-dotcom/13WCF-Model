# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

13WCF Data Loader - A tool to automate loading data from Settlement, Supplier, and Customer source files into the Bullish 13-Week Cash Flow (13WCF) Excel model.

**Two Implementations:**
1. **Browser-based** (`13WCF_Data_Loader_v3.html`) - Pure HTML/JavaScript using SheetJS
2. **Python script** (`13wcf_data_loader.py`) - Uses openpyxl library (recommended for large files)

## Running the Tool

**Python Version (Recommended):**
```bash
python3 13wcf_data_loader.py
```
The script will interactively prompt for file paths. You can drag and drop files into the terminal.

**Browser Version:**
Simply open `13WCF_Data_Loader_v3.html` in a modern browser (Chrome, Firefox, Edge, Safari).

## Required Input Files

The tool requires 4 files:

1. **13WCF Model** (`Bullish - 13WCF - DRAFT .xlsx`)
   - The master workbook that receives all data

2. **Settlement Details** (`Settlement_details.xlsx`)
   - Target sheet: `OPEX - Weekly Settlement Run`
   - Data columns: A-AR (1-44)
   - Formula columns: AT-BD (46-56)
   - Starting row: 9

3. **Supplier Invoices** (`Find_Supplier_Invoices_-_Devi.xlsx`)
   - Target sheet: `OPEX - AP`
   - Data columns: A-AM (1-39)
   - Formula columns: AO-BD (41-56)
   - Starting row: 13

4. **Customer Invoices** (`Customer_Invoice_Details.xlsx`)
   - Target sheet: `AR - Customer Invoice Detail`
   - Data columns: A-Q (1-17)
   - Formula columns: R-Y (18-25)
   - Starting row: 8

## Architecture

### Data Processing Flow

For each input file, the tool:
1. Opens the source 13WCF model
2. Locates the target worksheet
3. Saves formula templates from the first data row
4. Clears existing data (preserves headers)
5. Pastes new data from source file (skips first 2 header rows)
6. Copies and adjusts formula templates for each new row
7. Saves updated workbook

### Formula Handling

**Critical Requirement:** Formulas must be adjusted when copied to new rows.

The `adjust_formula()` function uses regex to update row references:
- Relative references (e.g., `A9`) are adjusted by the row offset
- Absolute row references (e.g., `A$9`) are NOT adjusted
- Pattern: `(\$?[A-Z]{1,3})(\$?)(\d+)`

Example:
```python
# Original formula in row 9:
=IF(A9="", "", VLOOKUP(A9, Table1, 2, FALSE))

# After pasting to row 10 (offset=1):
=IF(A10="", "", VLOOKUP(A10, Table1, 2, FALSE))
```

### Python Implementation Details

**openpyxl settings:**
- `data_only=True` for reading (gets calculated values, not formulas)
- `read_only=True` for source files (memory efficiency)
- Write mode for target workbook (allows modifications)

**Memory considerations:**
- Source files opened in read-only mode and closed after processing
- Progress logging every 500 rows
- Each sheet processed sequentially

### Browser Implementation Details (v3.0)

**SheetJS optimization flags:**
- `cellFormula: false` - Disables formula parsing to reduce memory
- `sheetStubs: false` - Skips empty cells
- Async breaks every 100 rows prevent UI freezing

**Known limitation:** May still crash on very large files (50MB+) due to browser memory constraints. Use Python version for large files.

## Development Notes

### When Modifying Column Ranges

If the 13WCF model structure changes, update these constants in both implementations:

**Settlement Run:**
- Data columns: 1-44 (A-AR)
- Formula columns: 46-56 (AT-BD)
- Start row: 9

**Supplier Invoices:**
- Data columns: 1-39 (A-AM)
- Formula columns: 41-56 (AO-BD)
- Start row: 13

**Customer Invoices:**
- Data columns: 1-17 (A-Q)
- Formula columns: 18-25 (R-Y)
- Start row: 8

### Testing

When testing changes:
1. Use small sample files first (< 100 rows)
2. Verify formula adjustment with spot checks
3. Compare row counts: input vs. output
4. Check that formulas reference correct rows
5. Verify no data in cleared columns (e.g., column AS for Settlement Run)

### Common Issues

**Empty rows in output:**
- Source files may have blank rows that get skipped (intentional behavior)
- Only rows with data in the data columns are copied

**Formula errors (#REF!):**
- Check that formula templates are captured correctly from the first data row
- Verify absolute vs. relative reference handling in `adjust_formula()`

**File size:**
- Python output files are typically larger than input due to unoptimized XML
- This is normal openpyxl behavior and doesn't affect Excel functionality
