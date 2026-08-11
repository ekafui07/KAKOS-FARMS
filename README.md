# KAKOS Audit Tool

A local, browser-based bank statement audit tool built with Flask. Upload a bank statement (CSV, Word, or PDF), and it parses transactions into a clean, filterable table with automatic Ghana Cedi (GH₵) formatting, KPI summaries, and one-click Excel export — all processed locally, nothing leaves your machine.

## Table of Contents

- [Features](#features)
- [Supported File Formats](#supported-file-formats)
- [Requirements](#requirements)
- [Installation](#installation)
- [Running the App](#running-the-app)
- [Usage](#usage)
- [Excel Export](#excel-export)
- [Notes on Parsing](#notes-on-parsing)

## Features

| Feature | Description |
|---|---|
| Multi-format parsing | Dedicated parser engines for CSV, DOCX, and PDF bank statements |
| Auto currency cleanup | Handles commas, parentheses (negatives), and `GH₵` symbols |
| Date/search filtering | Filter transactions by date range and free-text search across description and notes |
| KPI dashboard | Live totals for inflow, outflow, net movement, and closing balance |
| Excel export | Download a formatted `.xlsx` with a transaction sheet and a summary sheet |
| Local-only processing | No external services — statements are parsed and held in memory on your machine |

## Supported File Formats

- **`.csv`** — line-based parser using date-pattern detection to segment transaction blocks
- **`.docx`** — table-based parser for statements exported as Word documents, including multi-row/continuation entries
- **`.pdf`** — table-extraction parser (via `pdfplumber`) for statements exported as PDF, built for Universal Merchant Bank–style multi-column layouts

## Requirements

- Python 3.9+
- Dependencies listed in `requirements.txt` (Flask, pandas, pdfplumber, python-docx, xlsxwriter, and supporting libraries)

## Installation

```bash
git clone https://github.com/ekafui07/KAKOS-FARMS.git
cd KAKOS-FARMS
pip install -r requirements.txt
```

## Running the App

```bash
python kakos_audit.py
```

By default the app runs on port `5000` (override with the `PORT` environment variable). Open **http://localhost:5000** in your browser.

## Usage

1. Open the app and upload a bank statement (`.csv`, `.docx`, or `.pdf`)
2. Review the parsed transaction table and KPI summary (inflow, outflow, net movement, closing balance)
3. Use the **date range** and **search** filters to narrow down transactions
4. Click **Export Excel** to download a formatted spreadsheet of the current view
5. Use **Reset / Upload New** to clear the session and start over

## Excel Export

The exported `.xlsx` file includes:

- **Audit Data** sheet — the filtered transaction table with currency formatting, frozen header row, and autofilter enabled
- **Summary** sheet — total inflow, total outflow, net movement, closing balance, and transaction count for the current filter set

## Notes on Parsing

- Each file format has its own parser class (`BankParser`, `DocxBankParser`, `PdfBankParser`) tuned to that format's layout quirks
- Multi-line transaction notes (e.g. cheque references, continuation rows) are captured separately as **Extracted Notes**
- Uploads are capped at 20 MB
- If a file fails to parse or no transactions are found, the app surfaces an error instead of silently returning empty data
