# ROW-GIS-Pipeline

![Python](https://img.shields.io/badge/Python-3.14-blue?logo=python&logoColor=white)
![Status](https://img.shields.io/badge/Status-Active-brightgreen)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)

A general-purpose Python pipeline for **cross-file Excel intersection and comparison analysis**. Drop any Excel files into the `assets/` folder and instantly get a fully highlighted, color-coded, and annotated Excel report showing exactly which rows and columns match across all your files.

---

## Table of Contents

- [Overview](#overview)
- [Project Structure](#project-structure)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Output](#output)
- [Scripts](#scripts)
- [Contributing](#contributing)
- [Author](#author)

---

## Overview

This tool solves a common data problem:

> **"I have multiple Excel files with overlapping columns. Which rows share the same values across all of them — and exactly where?"**

Given 2 to 5 Excel files, the pipeline:
- Automatically finds all columns common to every file
- Finds every value that appears in ALL files under each common column
- Highlights matched cells in both the source files and result sheets using a unique color per column
- Adds comments directly on matched cells telling you exactly which file and row it matched with
- Compresses long row lists into readable ranges (e.g. `Row 2–44` instead of listing every row)
- Exports everything into a single clean Excel report

---

## Project Structure

```
ROW-GIS-Pipeline/
│
├── assets/                        # Place your input Excel files here (2–5 files)
│   ├── File1.xls or .xlsx
│   ├── File2.xls or .xlsx
│   └── ...
│
├── output/                        # Auto-created — all results written here
│   ├── comparison_output.xlsx     # Result from compare_spans_structures.py
│   └── intersection_output.xlsx   # Result from excel_intersect.py
│
├── compare_spans_structures.py    # Comparison tool for exactly 2 specific files
├── excel_intersect.py             # General intersection tool for 2–5 any files
├── requirements.txt               # Python dependencies
└── README.md                      # Project documentation
```

---

## Features

### `excel_intersect.py` — General Multi-File Intersection Tool
- Accepts **2 to 5 Excel files** of any name — no configuration needed
- Auto-detects all `.xls` / `.xlsx` files from the `assets/` folder
- **4-pass value normalization** before any comparison:
  - Convert to string → strip whitespace → remove trailing `.0` → lowercase
- **4-pass match verification** per value:
  - Set intersection → forward verify → reverse verify → cross verify
- **Hash map O(1) lookup** — fast and efficient even on large files
- **Smart row range compression:**
  `[2, 3, 4, 5, 30, 31]` → `"Row 2–5, Row 30–31"`
- **Every match labeled with file name** — never ambiguous:
  `Spans.xls: Row 2–44  matched with  Structures.xls: Row 5–20`
- **Color-coded highlights** — unique color per common column on all sheets
- **Cell comments** on every highlighted cell showing exactly what it matched with
- **Color legend** added below data on every sheet

### `compare_spans_structures.py` — Two-File Specific Comparison
- Dedicated comparison for exactly 2 named files
- Column-by-column cross-table matching
- `Match_Count`, `Matched_Rows`, and `Match_Detail` extra columns
- Highlighted matched cells with comments on both sheets
- Multi-sheet output with per-column breakdown

---

## Requirements

- Python 3.10 or higher
- Windows (tested on Windows 10/11)

### Python Dependencies

```
pandas
openpyxl
xlrd
```

---

## Installation

**1. Clone the repository:**
```bash
git clone https://github.com/your-username/ROW-GIS-Pipeline.git
cd ROW-GIS-Pipeline
```

**2. Install dependencies:**
```bash
C:\Python314\python.exe -m pip install pandas openpyxl xlrd
```

Or using the requirements file:
```bash
C:\Python314\python.exe -m pip install -r requirements.txt
```

---

## Usage

### General Multi-File Intersection (2–5 files)
1. Place **2 to 5** Excel files (`.xls` or `.xlsx`) inside the `assets/` folder
2. Run:
```bash
C:\Python314\python.exe excel_intersect.py
```

### Two-File Specific Comparison
1. Place your two Excel files inside the `assets/` folder
2. Run:
```bash
C:\Python314\python.exe compare_spans_structures.py
```

> **No configuration needed.** Both scripts auto-detect files from the `assets/` folder.

---

## Output

### `intersection_output.xlsx`

| Sheet | Contents |
|---|---|
| `Summary` | Per-column match counts, compressed row ranges per file, color reference |
| `Intersection_Data` | All matched rows from all files with full cross-file match detail |
| `[Column Name]` | One sheet per common column — all intersecting values with file ↔ file row mapping |
| `[File Name]` | Original file data with matched cells highlighted and commented |

### `comparison_output.xlsx`

| Sheet | Contents |
|---|---|
| `Summary` | Common columns overview with match statistics |
| `Intersection_Data` | All matched rows across both files |
| `[Column Name]` | One sheet per common column |
| `[File Name]` | Original data per file with highlights and comments |

### Highlight Color System

| Element | Color |
|---|---|
| Matched cell (common column 1) | 🟡 Yellow |
| Matched cell (common column 2) | 🟢 Green |
| Matched cell (common column 3) | 🔵 Blue |
| ... | Unique color per column |
| Extra result columns | 🔷 Light Blue |
| Header row | 🟦 Dark Blue |

### Cell Comment Format
Every highlighted matched cell contains a comment:
```
Column : Segment
Value  : SEG-C
This cell is in: Spans.xls Row 27
Matched in:
  Structures.xls: Row 5–20
  File3.xlsx: Row 8–14
```

---

## Scripts

| Script | Purpose | Input | Output |
|---|---|---|---|
| `excel_intersect.py` | General intersection for 2–5 any Excel files | `assets/*.xls` / `*.xlsx` | `output/intersection_output.xlsx` |
| `compare_spans_structures.py` | Dedicated comparison for 2 specific files | `assets/*.xls` / `*.xlsx` | `output/comparison_output.xlsx` |

---

## Contributing

1. Fork the repository
2. Create a feature branch:
```bash
git checkout -b feature/your-feature-name
```
3. Commit your changes:
```bash
git commit -m "Add: your feature description"
```
4. Push to the branch:
```bash
git push origin feature/your-feature-name
```
5. Open a Pull Request

---

## Author

**Jeeban Bashyal**
Computer Science — Alabama A&M University
NSBE Collegiate Member — Region 3, Chapter AAMU
