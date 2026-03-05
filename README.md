# 13WCF Data Loader

A browser-based data loader tool for the Bullish 13-Week Cash Flow (13WCF) Model.

## Overview

This tool automatically loads data from multiple source files into the 13WCF Excel model, eliminating manual copy-paste operations and ensuring data accuracy.

## Features

- **Browser-based**: No installation required - runs entirely in your web browser
- **Drag & Drop**: Simple file upload interface
- **Automated Processing**: Loads data into 3 worksheets automatically:
  - OPEX - Weekly Settlement Run
  - OPEX - AP (Supplier Invoices)
  - AR - Customer Invoice Detail
- **Formula Preservation**: Maintains all Excel formulas with proper row adjustments
- **Professional UI**: Bullish-branded interface with real-time processing feedback

## Required Files

### Source File (Required)
- **13WCF Model**: Bullish - 13WCF - DRAFT .xlsx
  - The master workbook that will receive all input data

### Input Files (All Required)
1. **Settlement Details**: Settlement_details.xlsx
   - Target: OPEX - Weekly Settlement Run tab
   - Data columns: A-AR (columns 1-44)
   - Formula columns: AT-BD (columns 46-56)

2. **Supplier Invoices**: Find_Supplier_Invoices_-_Devi.xlsx
   - Target: OPEX - AP tab
   - Data columns: A-AM (columns 1-39)
   - Formula columns: AO-BD (columns 41-56)

3. **Customer Invoices**: Customer_Invoice_Details.xlsx
   - Target: AR - Customer Invoice Detail tab
   - Data columns: A-Q (columns 1-17)
   - Formula columns: R-Y (columns 18-25)

## How It Works

1. **Upload Files**: Drag and drop or click to upload the 4 required files
2. **Process**: Click "Load Data into 13WCF" button
3. **Download**: Download the updated 13WCF model with all data loaded

### Processing Details

For each input file:
- Clears existing data from target worksheet (preserves headers)
- Pastes new data starting from the appropriate row
- Copies formula templates from the first data row
- Adjusts formula row references automatically
- Updates worksheet ranges to match new data extent

## Technical Details

- **Technology**: Pure HTML/JavaScript using SheetJS (xlsx.js)
- **File Format**: Excel .xlsx files (Office Open XML)
- **Formula Handling**: Preserves and adjusts cell formulas
- **VBA Support**: Maintains VBA macros in the output file
- **Client-Side Only**: All processing happens in your browser (no data uploaded to servers)

## Usage

Simply open `13WCF_Data_Loader.html` in any modern web browser (Chrome, Firefox, Edge, Safari).

## Version

Current Version: 1.0.0
Last Updated: 2026-03-05

## Department

Bullish Treasury & FP&A

---

© 2026 Bullish. All rights reserved.
