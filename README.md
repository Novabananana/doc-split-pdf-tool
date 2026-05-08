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

### 📑 Tab 1: Split by ID
- Group all pages with the same ID into a single PDF
- Extract unlimited custom data fields using text prefixes
- Generate CSV reports with all extracted data
- Multi-criteria file naming (up to 4 fields)
- Custom filename separator and suffix

### 🔍 Tab 2: Split by Name  
- Process PDFs without ID numbers (certificates, testamurs)
- Match names to IDs using CSV mapping
- Automatically loads Tab 1 CSV as default mapping
- Multi-line extraction with optional "Stop Text" feature

### ✂️ Tab 3: Split by Page Range
- Extract specific page ranges from PDFs
- Support for multiple ranges (up to 20)
- CSV output with extraction details

### 🎨 Interface
- Clean Windows 10 flat design
- Dismissible tips for new users


## 🚀 Quick Start

### Windows (Recommended)
1. Download `DocSplit.exe` from [Releases](https://github.com/Novabananana/doc-split-pdf-tool/releases)
2. Run the executable (no installation required)

### From Source
```bash
# Clone the repository
git clone https://github.com/Novabananana/doc-split-pdf-tool.git
cd doc-split-pdf-tool

# Install dependencies
pip install -r requirements.txt

# Run the application
python docu_split.py
📖 Usage Guide
Basic Workflow
Tab 1 (Documents with IDs):

Select your PDF file

Choose output folder

Add extraction criteria (e.g., "Student Number:", "Student Name:")

Select which fields to include in filenames

Click "SPLIT BY ID"

Tab 2 (Documents without IDs):

Process Tab 1 first to create CSV mapping

Load CSV (auto-detected from Tab 1)

Configure Tab 2 extraction criteria

Click "SPLIT BY NAME"

Tab 3 (Page Ranges):

Select PDF file

Enter page ranges (e.g., "1-5, 10-15, 20")

Choose output folder

Click "EXTRACT PAGES"

Advanced Features
Stop Text - Stop reading at a specific phrase for multi-line values:

Prefix: "Degree of"

Stop Text: "on the day of"

Result: Captures only the degree name without the date

Custom Filename - Combine up to 4 criteria with custom separator:

Example: Document ID_Name_Description.pdf

Separator options: _ -  . __

Optional suffix: _final → Document ID_Name_final.pdf

🛠️ Building from Source
bash
# Install PyInstaller
pip install pyinstaller

# Build executable
pyinstaller --onefile --name "DocSplit" --windowed --hidden-import PyPDF2 --hidden-import fitz docu_split.py
📋 Requirements
Windows 10/11 (Linux/Mac supported from source)

Python 3.7+ (for source installation)

4GB RAM recommended

100MB disk space

🤝 Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

Fork the repository

Create your feature branch (git checkout -b feature/AmazingFeature)

Commit your changes (git commit -m 'Add some AmazingFeature')

Push to the branch (git push origin feature/AmazingFeature)

Open a Pull Request

📝 Changelog
v1.0.0 (2026-05-08)
Initial release

Split by ID, Name, and page ranges

Custom extraction criteria

CSV output support

Dark/Light mode

Dismissible tips

⚠️ Known Issues
None currently

🙏 Acknowledgments
PyPDF2 - PDF manipulation

PyMuPDF - PDF text extraction

Pillow - Image handling

📧 Contact
Author: Novabananana

GitHub: @Novabananana

Issues: Report a bug

📜 License
Distributed under the GNU General Public License v3.0. See LICENSE for more information.

⭐ Star this repo if you find it useful!

Made with ❤️ for the PDF splitting community

text

### Step 4: Paste into nano

- Right-click in the terminal (or press `Ctrl + Shift + V`) to paste
- Or use `Shift + Insert` to paste

### Step 5: Save and exit

- Press `Ctrl + O` (save)
- Press `Enter` (confirm filename)
- Press `Ctrl + X` (exit)

### Step 6: Commit and push the changes

```bash
git add README.md
git commit -m "Improve README with detailed documentation"
git push
