"""
13WCF Data Loader - Documentation Dashboard Server
===================================================
A Flask web server that provides an interactive documentation dashboard.

Usage:
    python3 documentation_server.py

Then open your browser to: http://localhost:5001
"""

from flask import Flask, render_template_string
import os
import webbrowser
from threading import Timer

app = Flask(__name__)

# HTML template for the dashboard
DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>13WCF Data Loader - Documentation Dashboard</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .dashboard-container {
            max-width: 1400px;
            margin: 0 auto;
        }

        .header {
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }

        .header h1 {
            color: #667eea;
            font-size: 2.5em;
            margin-bottom: 10px;
        }

        .header p {
            color: #666;
            font-size: 1.1em;
        }

        .main-grid {
            display: grid;
            grid-template-columns: 300px 1fr;
            gap: 30px;
        }

        .sidebar {
            background: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            height: fit-content;
            position: sticky;
            top: 20px;
        }

        .sidebar h3 {
            color: #667eea;
            margin-bottom: 20px;
            font-size: 1.3em;
        }

        .nav-item {
            padding: 12px 15px;
            margin-bottom: 8px;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.3s ease;
            color: #333;
            font-weight: 500;
        }

        .nav-item:hover {
            background: #f0f0f0;
            transform: translateX(5px);
        }

        .nav-item.active {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }

        .content-area {
            background: white;
            padding: 40px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            min-height: 600px;
        }

        .section {
            display: none;
        }

        .section.active {
            display: block;
            animation: fadeIn 0.5s;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .section h2 {
            color: #667eea;
            font-size: 2em;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #667eea;
        }

        .section h3 {
            color: #764ba2;
            font-size: 1.5em;
            margin-top: 30px;
            margin-bottom: 15px;
        }

        .section h4 {
            color: #333;
            font-size: 1.2em;
            margin-top: 20px;
            margin-bottom: 10px;
        }

        .section p {
            line-height: 1.8;
            color: #555;
            margin-bottom: 15px;
        }

        .section ul, .section ol {
            margin-left: 30px;
            margin-bottom: 20px;
            line-height: 1.8;
            color: #555;
        }

        .section li {
            margin-bottom: 8px;
        }

        .code-block {
            background: #f8f9fa;
            border-left: 4px solid #667eea;
            padding: 20px;
            border-radius: 8px;
            margin: 20px 0;
            font-family: 'Monaco', 'Courier New', monospace;
            overflow-x: auto;
            white-space: pre;
            font-size: 0.9em;
            line-height: 1.6;
        }

        .inline-code {
            background: #f0f0f0;
            padding: 3px 8px;
            border-radius: 4px;
            font-family: 'Monaco', 'Courier New', monospace;
            font-size: 0.9em;
            color: #d63384;
        }

        .info-box {
            background: #e7f3ff;
            border-left: 4px solid #2196F3;
            padding: 20px;
            margin: 20px 0;
            border-radius: 8px;
        }

        .warning-box {
            background: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 20px;
            margin: 20px 0;
            border-radius: 8px;
        }

        .file-card {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin: 15px 0;
            border: 2px solid #e0e0e0;
        }

        .file-card h4 {
            color: #667eea;
            margin-top: 0;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }

        table th, table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #e0e0e0;
        }

        table th {
            background: #f8f9fa;
            color: #667eea;
            font-weight: 600;
        }

        .badge {
            display: inline-block;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
            margin-right: 10px;
        }

        .badge-primary {
            background: #667eea;
            color: white;
        }

        .badge-secondary {
            background: #764ba2;
            color: white;
        }

        .badge-success {
            background: #28a745;
            color: white;
        }
    </style>
</head>
<body>
    <div class="dashboard-container">
        <div class="header">
            <h1>📊 13WCF Data Loader Documentation</h1>
            <p>Interactive guide for the Bullish 13-Week Cash Flow data loading tool</p>
        </div>

        <div class="main-grid">
            <div class="sidebar">
                <h3>Navigation</h3>
                <div class="nav-item active" onclick="showSection('overview')">📖 Overview</div>
                <div class="nav-item" onclick="showSection('quickstart')">🚀 Quick Start</div>
                <div class="nav-item" onclick="showSection('files')">📁 Required Files</div>
                <div class="nav-item" onclick="showSection('architecture')">🏗️ Architecture</div>
                <div class="nav-item" onclick="showSection('formulas')">🔢 Formula Handling</div>
                <div class="nav-item" onclick="showSection('python')">🐍 Python Details</div>
                <div class="nav-item" onclick="showSection('browser')">🌐 Browser Details</div>
                <div class="nav-item" onclick="showSection('development')">⚙️ Development</div>
                <div class="nav-item" onclick="showSection('troubleshooting')">🔧 Troubleshooting</div>
            </div>

            <div class="content-area">
                <!-- Overview Section -->
                <div id="overview" class="section active">
                    <h2>Project Overview</h2>
                    <p>The <strong>13WCF Data Loader</strong> automates the process of loading data from Settlement, Supplier, and Customer source files into the Bullish 13-Week Cash Flow (13WCF) Excel model.</p>

                    <div class="info-box">
                        <strong>💡 Key Benefit:</strong> Eliminates manual copy-paste operations and ensures data accuracy through automated formula adjustment.
                    </div>

                    <h3>Two Implementations Available</h3>

                    <div class="file-card">
                        <h4>🐍 Python Script (Recommended)</h4>
                        <p><span class="inline-code">13wcf_data_loader.py</span></p>
                        <p>Uses openpyxl library for efficient handling of large files. Ideal for files over 50MB.</p>
                        <span class="badge badge-success">Recommended</span>
                        <span class="badge badge-primary">Large Files</span>
                    </div>

                    <div class="file-card">
                        <h4>🌐 Browser Version</h4>
                        <p><span class="inline-code">13WCF_Data_Loader_v3.html</span></p>
                        <p>Pure HTML/JavaScript using SheetJS. No installation required, runs in your browser.</p>
                        <span class="badge badge-secondary">No Install</span>
                    </div>

                    <h3>What It Does</h3>
                    <p>The tool processes <strong>3 data sources</strong> into <strong>3 worksheets</strong>:</p>
                    <table>
                        <thead>
                            <tr>
                                <th>Source File</th>
                                <th>Target Worksheet</th>
                                <th>Function</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>Settlement Details</td>
                                <td>OPEX - Weekly Settlement Run</td>
                                <td>Weekly settlement data</td>
                            </tr>
                            <tr>
                                <td>Supplier Invoices</td>
                                <td>OPEX - AP</td>
                                <td>Accounts payable data</td>
                            </tr>
                            <tr>
                                <td>Customer Invoices</td>
                                <td>AR - Customer Invoice Detail</td>
                                <td>Accounts receivable data</td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <!-- Quick Start Section -->
                <div id="quickstart" class="section">
                    <h2>Quick Start Guide</h2>

                    <h3>Python Version (Recommended)</h3>
                    <div class="code-block">python3 13wcf_data_loader.py</div>
                    <p>The script will interactively prompt you for each file. You can drag and drop files into the terminal window.</p>

                    <h3>Browser Version</h3>
                    <ol>
                        <li>Open <span class="inline-code">13WCF_Data_Loader_v3.html</span> in Chrome, Firefox, Edge, or Safari</li>
                        <li>Drag and drop the 4 required files into the upload areas</li>
                        <li>Click "Load Data into 13WCF"</li>
                        <li>Download the updated file when processing completes</li>
                    </ol>

                    <div class="warning-box">
                        <strong>⚠️ Browser Limitation:</strong> The browser version may crash on very large files (50MB+) due to memory constraints. Use the Python version for large files.
                    </div>

                    <h3>Output</h3>
                    <p><strong>Python:</strong> Creates <span class="inline-code">Bullish_13WCF_Updated_YYYYMMDD_HHMMSS.xlsx</span> in the same directory</p>
                    <p><strong>Browser:</strong> Downloads the updated file automatically</p>
                </div>

                <!-- Required Files Section -->
                <div id="files" class="section">
                    <h2>Required Input Files</h2>
                    <p>The tool requires <strong>4 files total</strong>: 1 master workbook + 3 data source files.</p>

                    <div class="file-card">
                        <h4>1. 13WCF Model (Master Workbook)</h4>
                        <p><strong>Filename:</strong> <span class="inline-code">Bullish - 13WCF - DRAFT .xlsx</span></p>
                        <p>The master workbook that will receive all input data.</p>
                    </div>

                    <div class="file-card">
                        <h4>2. Settlement Details</h4>
                        <p><strong>Filename:</strong> <span class="inline-code">Settlement_details.xlsx</span></p>
                        <p><strong>Target Sheet:</strong> OPEX - Weekly Settlement Run</p>
                        <table>
                            <tr>
                                <td><strong>Data Columns:</strong></td>
                                <td>A-AR (columns 1-44)</td>
                            </tr>
                            <tr>
                                <td><strong>Formula Columns:</strong></td>
                                <td>AT-BD (columns 46-56)</td>
                            </tr>
                            <tr>
                                <td><strong>Starting Row:</strong></td>
                                <td>9</td>
                            </tr>
                        </table>
                    </div>

                    <div class="file-card">
                        <h4>3. Supplier Invoices</h4>
                        <p><strong>Filename:</strong> <span class="inline-code">Find_Supplier_Invoices_-_Devi.xlsx</span></p>
                        <p><strong>Target Sheet:</strong> OPEX - AP</p>
                        <table>
                            <tr>
                                <td><strong>Data Columns:</strong></td>
                                <td>A-AM (columns 1-39)</td>
                            </tr>
                            <tr>
                                <td><strong>Formula Columns:</strong></td>
                                <td>AO-BD (columns 41-56)</td>
                            </tr>
                            <tr>
                                <td><strong>Starting Row:</strong></td>
                                <td>13</td>
                            </tr>
                        </table>
                    </div>

                    <div class="file-card">
                        <h4>4. Customer Invoices</h4>
                        <p><strong>Filename:</strong> <span class="inline-code">Customer_Invoice_Details.xlsx</span></p>
                        <p><strong>Target Sheet:</strong> AR - Customer Invoice Detail</p>
                        <table>
                            <tr>
                                <td><strong>Data Columns:</strong></td>
                                <td>A-Q (columns 1-17)</td>
                            </tr>
                            <tr>
                                <td><strong>Formula Columns:</strong></td>
                                <td>R-Y (columns 18-25)</td>
                            </tr>
                            <tr>
                                <td><strong>Starting Row:</strong></td>
                                <td>8</td>
                            </tr>
                        </table>
                    </div>
                </div>

                <!-- Architecture Section -->
                <div id="architecture" class="section">
                    <h2>Architecture & Data Flow</h2>

                    <h3>Processing Flow</h3>
                    <p>For each input file, the tool follows these steps:</p>
                    <ol>
                        <li><strong>Open</strong> the source 13WCF model workbook</li>
                        <li><strong>Locate</strong> the target worksheet by name</li>
                        <li><strong>Save</strong> formula templates from the first data row</li>
                        <li><strong>Clear</strong> existing data (preserves headers)</li>
                        <li><strong>Paste</strong> new data from source file (skips first 2 header rows)</li>
                        <li><strong>Copy</strong> and adjust formula templates for each new row</li>
                        <li><strong>Save</strong> the updated workbook</li>
                    </ol>

                    <div class="info-box">
                        <strong>💡 Key Feature:</strong> Empty rows in source data are automatically skipped. Only rows with data are copied.
                    </div>

                    <h3>Column Mapping</h3>
                    <p>Each worksheet has two types of columns:</p>
                    <ul>
                        <li><strong>Data Columns:</strong> Values copied directly from source file</li>
                        <li><strong>Formula Columns:</strong> Formulas that reference the data columns, automatically adjusted for each row</li>
                    </ul>

                    <div class="warning-box">
                        <strong>⚠️ Important:</strong> The gap between data and formula columns (e.g., column AS for Settlement) is intentionally left empty. This is by design.
                    </div>
                </div>

                <!-- Formulas Section -->
                <div id="formulas" class="section">
                    <h2>Formula Handling</h2>

                    <h3>The Challenge</h3>
                    <p>When copying formulas to new rows, cell references must be adjusted to point to the correct row. This is critical for maintaining data integrity.</p>

                    <h3>How It Works</h3>
                    <p>The <span class="inline-code">adjust_formula()</span> function uses regex pattern matching to identify and adjust cell references:</p>

                    <div class="code-block">Pattern: (\\$?[A-Z]{1,3})(\\$?)(\\d+)

Matches:
- Column letter(s) with optional $ prefix
- Optional $ before row number
- Row number</div>

                    <h3>Adjustment Rules</h3>
                    <table>
                        <thead>
                            <tr>
                                <th>Reference Type</th>
                                <th>Example</th>
                                <th>Action</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>Relative</td>
                                <td><span class="inline-code">A9</span></td>
                                <td>✅ Adjusted by row offset</td>
                            </tr>
                            <tr>
                                <td>Absolute Row</td>
                                <td><span class="inline-code">A$9</span></td>
                                <td>❌ NOT adjusted</td>
                            </tr>
                            <tr>
                                <td>Absolute Column</td>
                                <td><span class="inline-code">$A9</span></td>
                                <td>✅ Row adjusted, column stays absolute</td>
                            </tr>
                        </tbody>
                    </table>

                    <h3>Example</h3>
                    <div class="code-block">Original formula (row 9):
=IF(A9="", "", VLOOKUP(A9, Table1, 2, FALSE))

After copying to row 10 (offset = 1):
=IF(A10="", "", VLOOKUP(A10, Table1, 2, FALSE))

After copying to row 100 (offset = 91):
=IF(A100="", "", VLOOKUP(A100, Table1, 2, FALSE))</div>

                    <div class="info-box">
                        <strong>💡 Why This Matters:</strong> Without proper adjustment, formulas would reference the wrong rows, producing incorrect calculations throughout the entire model.
                    </div>
                </div>

                <!-- Python Details Section -->
                <div id="python" class="section">
                    <h2>Python Implementation Details</h2>

                    <h3>Technology Stack</h3>
                    <ul>
                        <li><strong>openpyxl:</strong> Excel file manipulation library</li>
                        <li><strong>Python 3.x:</strong> No external dependencies beyond openpyxl</li>
                    </ul>

                    <h3>openpyxl Settings</h3>
                    <div class="code-block">For reading source files:
- data_only=True    # Gets calculated values, not formulas
- read_only=True    # Memory efficiency

For writing target workbook:
- Default mode      # Allows modifications</div>

                    <h3>Memory Optimization</h3>
                    <ul>
                        <li>Source files opened in <strong>read-only mode</strong> and closed immediately after processing</li>
                        <li>Progress logging every <strong>500 rows</strong> to track large file processing</li>
                        <li>Each sheet processed <strong>sequentially</strong> to avoid loading all data at once</li>
                    </ul>

                    <h3>File Structure</h3>
                    <p>The Python script has three main processing functions:</p>
                    <ul>
                        <li><span class="inline-code">process_settlement_run()</span> - Handles Settlement Details</li>
                        <li><span class="inline-code">process_supplier_invoices()</span> - Handles Supplier Invoices</li>
                        <li><span class="inline-code">process_customer_invoices()</span> - Handles Customer Invoices</li>
                    </ul>

                    <h3>Output File Naming</h3>
                    <p>Output files include a timestamp for version tracking:</p>
                    <div class="code-block">Bullish_13WCF_Updated_20260305_143022.xlsx
                        ↑                    ↑        ↑
                    Prefix            Date    Time</div>
                </div>

                <!-- Browser Details Section -->
                <div id="browser" class="section">
                    <h2>Browser Implementation Details (v3.0)</h2>

                    <h3>Technology Stack</h3>
                    <ul>
                        <li><strong>SheetJS (xlsx.js):</strong> JavaScript library for Excel file processing</li>
                        <li><strong>Pure HTML/JavaScript:</strong> No server required, runs entirely client-side</li>
                    </ul>

                    <h3>Memory Optimizations</h3>
                    <p>Version 3.0 includes several optimizations to prevent browser crashes:</p>

                    <div class="code-block">cellFormula: false     // Disables formula parsing
sheetStubs: false       // Skips empty cells
Async breaks every 100 rows  // Prevents UI freezing</div>

                    <h3>Processing Flow</h3>
                    <ol>
                        <li>Files loaded using FileReader API</li>
                        <li>Parsed with SheetJS into JSON format</li>
                        <li>Data processed row by row with async breaks</li>
                        <li>Output generated and downloaded automatically</li>
                    </ol>

                    <div class="warning-box">
                        <strong>⚠️ Known Limitation:</strong> May still crash on very large files (50MB+) due to browser memory constraints. This is a browser limitation, not a bug in the code.
                        <br><br>
                        <strong>Solution:</strong> Use the Python version for files larger than 50MB.
                    </div>

                    <h3>Progress Feedback</h3>
                    <p>The browser version includes a visual progress bar that updates in real-time during processing.</p>
                </div>

                <!-- Development Section -->
                <div id="development" class="section">
                    <h2>Development Notes</h2>

                    <h3>Modifying Column Ranges</h3>
                    <p>If the 13WCF model structure changes, update these constants in both implementations:</p>

                    <div class="file-card">
                        <h4>Settlement Run</h4>
                        <ul>
                            <li>Data columns: 1-44 (A-AR)</li>
                            <li>Formula columns: 46-56 (AT-BD)</li>
                            <li>Start row: 9</li>
                        </ul>
                    </div>

                    <div class="file-card">
                        <h4>Supplier Invoices</h4>
                        <ul>
                            <li>Data columns: 1-39 (A-AM)</li>
                            <li>Formula columns: 41-56 (AO-BD)</li>
                            <li>Start row: 13</li>
                        </ul>
                    </div>

                    <div class="file-card">
                        <h4>Customer Invoices</h4>
                        <ul>
                            <li>Data columns: 1-17 (A-Q)</li>
                            <li>Formula columns: 18-25 (R-Y)</li>
                            <li>Start row: 8</li>
                        </ul>
                    </div>

                    <h3>Testing Checklist</h3>
                    <p>When testing changes to the code:</p>
                    <ol>
                        <li>✅ Use small sample files first (&lt; 100 rows)</li>
                        <li>✅ Verify formula adjustment with spot checks</li>
                        <li>✅ Compare row counts: input vs. output</li>
                        <li>✅ Check that formulas reference correct rows</li>
                        <li>✅ Verify no data in cleared columns (e.g., column AS for Settlement Run)</li>
                    </ol>
                </div>

                <!-- Troubleshooting Section -->
                <div id="troubleshooting" class="section">
                    <h2>Troubleshooting</h2>

                    <h3>Empty Rows in Output</h3>
                    <p><strong>Issue:</strong> Output has fewer rows than source file</p>
                    <p><strong>Cause:</strong> Source files may have blank rows that get skipped</p>
                    <p><strong>Solution:</strong> This is intentional behavior. Only rows with data in the data columns are copied.</p>

                    <h3>Formula Errors (#REF!)</h3>
                    <p><strong>Issue:</strong> Formulas show #REF! errors after loading</p>
                    <p><strong>Possible Causes:</strong></p>
                    <ul>
                        <li>Formula templates not captured correctly from first data row</li>
                        <li>Absolute vs. relative reference handling issue in <span class="inline-code">adjust_formula()</span></li>
                    </ul>
                    <p><strong>Solution:</strong> Check that formula templates in the source model use relative references (e.g., <span class="inline-code">A9</span> not <span class="inline-code">A$9</span>)</p>

                    <h3>Large Output File Size</h3>
                    <p><strong>Issue:</strong> Python output files are larger than expected</p>
                    <p><strong>Cause:</strong> openpyxl doesn't optimize XML as aggressively as Excel does</p>
                    <p><strong>Solution:</strong> This is normal openpyxl behavior and doesn't affect Excel functionality. Files will compress to normal size when saved in Excel.</p>

                    <h3>Browser Crash ("Aw Snap")</h3>
                    <p><strong>Issue:</strong> Chrome crashes when processing large files</p>
                    <p><strong>Cause:</strong> Browser memory limitations with files over 50MB</p>
                    <p><strong>Solution:</strong> Use the Python version (<span class="inline-code">13wcf_data_loader.py</span>) for large files</p>

                    <h3>Sheet Not Found Error</h3>
                    <p><strong>Issue:</strong> Error message about missing sheet</p>
                    <p><strong>Cause:</strong> Source model doesn't have expected sheet names</p>
                    <p><strong>Solution:</strong> Verify the source model has these exact sheet names:</p>
                    <ul>
                        <li>OPEX - Weekly Settlement Run</li>
                        <li>OPEX - AP</li>
                        <li>AR - Customer Invoice Detail</li>
                    </ul>

                    <div class="info-box">
                        <strong>💡 Need More Help?</strong> Check the CLAUDE.md file for additional technical details, or review the source code comments for inline documentation.
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        function showSection(sectionId) {
            // Hide all sections
            const sections = document.querySelectorAll('.section');
            sections.forEach(section => section.classList.remove('active'));

            // Remove active from all nav items
            const navItems = document.querySelectorAll('.nav-item');
            navItems.forEach(item => item.classList.remove('active'));

            // Show selected section
            document.getElementById(sectionId).classList.add('active');

            // Highlight selected nav item
            event.target.classList.add('active');

            // Scroll to top of content area
            document.querySelector('.content-area').scrollTop = 0;
        }
    </script>
</body>
</html>
"""


@app.route('/')
def dashboard():
    """Serve the documentation dashboard"""
    return render_template_string(DASHBOARD_HTML)


def open_browser():
    """Open the browser after a short delay"""
    webbrowser.open('http://localhost:5001')


if __name__ == '__main__':
    print("\n" + "="*60)
    print("13WCF Data Loader - Documentation Dashboard")
    print("="*60)
    print("\nStarting server...")
    print("Dashboard will be available at: http://localhost:5001")
    print("\nPress Ctrl+C to stop the server")
    print("="*60 + "\n")

    # Open browser after 1 second delay
    Timer(1, open_browser).start()

    # Run Flask server
    app.run(host='localhost', port=5001, debug=False)
