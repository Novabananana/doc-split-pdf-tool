## 💝 Support the Project

If Doc Split saves you time and makes your work easier, consider supporting its development!

[![Support via Ko-fi](https://img.shields.io/badge/Support%20Me-Ko--fi-FF5E5B?style=for-the-badge&logo=ko-fi&logoColor=white)](https://ko-fi.com/novabananana)

Your support helps me add new features, fix bugs, and keep the project alive! 🙏


# 📄 Doc Split - PDF Document Splitter

[![Version](https://img.shields.io/badge/version-1.0.0-blue)](https://github.com/Novabananana/doc-split-pdf-tool/releases)
[![License](https://img.shields.io/badge/license-GPLv3-green)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.7+-blue)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20Mac-lightgrey)]()

A powerful desktop application for splitting PDF documents and extracting data with customizable criteria. Perfect for processing student records, certificates, invoices, or any multi-page PDF documents.


## ✨ Features

### Tab 1: Split by ID / Multi-Criteria
- **Smart Grouping** - Groups pages with same ID even if they're scattered throughout the document
- **Multi-Criteria Grouping** - Group by any extracted field (not just ID)
- **CSV Export** - Extract data to CSV with page information
- **CSV-Only Mode** - Extract data without generating PDF files
- **No Grouping Mode** - Export each page as a separate CSV row
- **Custom Extraction** - Configure unlimited extraction criteria with regex patterns
- **Stop Text** - Stop reading at specific phrases for multi-line values
- **Flexible Naming** - Use up to 4 extracted fields for output filenames

### Tab 2: Split by Name
- **CSV Mapping** - Match names from PDF to IDs from a CSV file
- **Smart Matching** - Direct, case-insensitive, and partial name matching
- **Batch Processing** - Process entire documents page by page

### Tab 3: Split by Page Range
- **Multiple Ranges** - Extract multiple page ranges at once (e.g., `1-5,10-15,20`)
- **Flexible Syntax** - Supports single pages, ranges, or mixed (`1,3-5,10`)
- **CSV Logging** - Track which pages were extracted from which ranges

### Tab 4: PDF Merger
- **Merge Multiple PDFs** - Combine several PDF files into one
- **Page Range Selection** - Include only specific pages from each file
- **Reorder Files** - Control merge order with up/down buttons
- **Table of Contents** - Add an automatic TOC page (requires reportlab)
- **Bookmarks** - Create outline bookmarks for each merged file

### General Features
- 🎨 **Modern Flat UI** - Clean Windows 10-style interface
- 💾 **Auto-Save Settings** - All preferences persist between sessions
- 📊 **CSV Output** - Extracted data saved with full metadata
- 🔄 **Multi-threaded** - Responsive UI during processing
- 🌙 **Light Theme Only** - Clean, professional appearance
- 📝 **Detailed Logging** - See exactly what's happening
- 🚀 **100% Offline** - No data leaves your computer


🚀 Quick Start
Example 1: Extract data from invoices
Open Tab 1

Click "Browse" to select your PDF

Configure criteria:

Add "Invoice Number" with prefix "Invoice #:"

Add "Customer Name" with prefix "Customer:"

Add "Amount" with prefix "Total:"

Select which fields to include in filenames

Click "SPLIT BY ID"

Example 2: Merge multiple reports
Go to Tab 4 (PDF Merger)

Click "Add PDF(s)" and select your files

Double-click any file to set page ranges

Use Move Up/Down to arrange order

Choose output folder and filename

Click "MERGE PDFS"

📖 Detailed Usage
Tab 1: Split by ID / Multi-Criteria
Setting Up Extraction Criteria
Add Criterion: Go to Criteria → Tab 1: Add Criterion

Configure:

Display Name: What to call this field (e.g., "Document ID")

Text Prefix: The exact text before your value (e.g., "ID Number:")

Stop Text: (Optional) Stop reading when this text appears

Data Type: "ID" for grouping, "Text" for regular fields

Grouping Options
Single Criterion (ID) : Groups pages by the ID field

Multiple Criteria (Custom) : Groups pages by any field you select

Export Modes
Grouped by Key : Pages with same key become one PDF/CSV row

Each Page Separately : Every page becomes its own CSV row (no PDFs)

CSV-Only Mode : Extract data without creating PDF files

Tab 2: Split by Name
Workflow
First process documents in Tab 1 to create a CSV file

Switch to Tab 2 and select your PDF

Load the CSV mapping file (auto-detected from Tab 1)

Configure extraction criteria (matches the PDF structure)

Click split - each page becomes a PDF named after matched IDs

Matching Logic
Direct match: "John Smith" → "John Smith"

Case-insensitive: "JOHN SMITH" → "John Smith"

Partial match: "J. Smith" → "John Smith" (if close)

Tab 3: Split by Page Range
Syntax Examples
5-10 - Pages 5 through 10

1,3,5 - Pages 1, 3, and 5

1-5,10-15,20 - Multiple ranges (max 20)

Tab 4: PDF Merger
Page Range Filters
Leave blank: Include all pages

1-5: Include only pages 1-5

3,7,10: Include specific pages

1-5,8,10-15: Mixed ranges

Features
Table of Contents: Adds a TOC page at the beginning

Bookmarks: Creates clickable outline entries for navigation

Live Preview: See file order and page counts

⚙️ Configuration
Settings Saved Automatically
Extraction criteria (separate for Tab 1 & 2)

File naming rules (selected fields, separator, suffix)

CSV output settings

Grouping preferences

Export mode selections

Window size and position

Menu Options
Menu	Action
File	Exit application
Criteria	Add/remove/reset extraction rules for Tab 1 and 2
View	Restore informational tips
Settings	Customize button text, rename tabs, save preferences
Help	Quick start guide and about information
💡 Tips & Tricks
Optimizing Extraction
Use Stop Text for multi-line values:

text
Description: This is a long description
that spans multiple lines
on the day of the event

Set Stop Text to "on the day of" to capture only the description
Test with CSV-Only Mode before generating PDFs

Use No Grouping Mode to quickly audit all pages

Check Page Numbers in CSV when debugging grouping issues

Best Practices
Always test with a small PDF first

Save criteria as they're auto-saved between sessions

Use descriptive names for extraction criteria

Keep CSV files for future reference and Tab 2 processing

Back up settings by saving the .doc_split_settings.json file

🔧 Troubleshooting
Common Issues
Issue	Solution
No IDs found	Check prefix matches exactly (spaces, colons)
Pages not grouping	Ensure ID field type is set to "ID"
CSV not loading	Check column names contain "name" or "id"
Merger TOC missing	Install reportlab: pip install reportlab
Large file slow	Use CSV-only mode for data extraction
Debugging Tools
Debug Tab 1 Criteria button - Shows current configuration

Debug PDF Text button - Displays raw text from first pages

Debug CSV button - Analyzes CSV structure for mapping

Processing Log - Real-time operation details

Getting Help
Check the Quick Start Guide (Help menu)

Review the processing log for error details

Use debug buttons to verify configurations

Test with a small sample PDF first

📄 License
This project is licensed under the MIT License - see the LICENSE file for details.

🙏 Acknowledgments
Built with PyPDF2 for PDF manipulation

Uses PyMuPDF (fitz) for text extraction

UI designed with Python's tkinter

<div align="center">
Made with ❤️ for efficient PDF document processing

Report Bug · Request Feature

</div> ```
