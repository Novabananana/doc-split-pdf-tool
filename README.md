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
- Dark/Light mode toggle
- Dismissible tips for new users
- Real-time filename preview

## 📸 Screenshot

![Doc Split Screenshot](screenshot.png)

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
