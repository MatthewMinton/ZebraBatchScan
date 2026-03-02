
# DS3678 Batch Scan → Text → Excel Logger

![Python](https://img.shields.io/badge/Python-3.14.3-blue)
![OS](https://img.shields.io/badge/OS-Windows%2010%20%7C%2011-green)
![Scanner](https://img.shields.io/badge/Scanner-Zebra%20DS3678-orange)
![Status](https://img.shields.io/badge/Status-Production%20Ready-brightgreen)

---

## Overview

This repository provides a **two-script Python solution** for logging **Zebra DS3678 barcode scanner batch scans** into a persistent text file and then appending that data into an Excel workbook.

It is designed for **manufacturing, controls, and traceability environments** where:

- Operators scan serial numbers offline using **batch (store-and-forward) mode**
- Scans are uploaded when the scanner is docked
- Data must be validated, retained, and archived
- Duplicate ingestion must be avoided
- Excel is the final reporting or handoff format

This solution intentionally **does not rely on MES integration** and is safe, auditable, and easy to maintain.

---

## Tested Environment

- **Python:** 3.14.3 (Windows)
- **Scanner:** Zebra DS3678
- **Connection:** USB cradle in HID Keyboard mode
- **Operating System:** Windows 10 / Windows 11

---

## High-Level Architecture

DS3678 Scanner (Batch Mode)
↓
USB Cradle (HID Keyboard)
↓
ZebraScanner.py
↓
Append-Only Text File (psion.txt)
↓
SerialToExcel.py (Scheduled Daily)
↓
Excel Log File (psion_log.xlsx)


---

## Features

### Scanner & Validation
- Supports **DS3678 batch (store-and-forward) uploads**
- Accepts **numeric-only serial numbers**
- Enforces **exactly 18 digits**
- Automatically splits **36-digit scans into two 18-digit serial numbers**
- Deduplicates serial numbers **within each batch**

### Text Logging
- Appends data to a **plain text file**
- Never overwrites existing data
- One record per line
- One timestamp per batch (intentional)
- Format:

MM/DD/YYYY,HH:MM:SS,Serial Number


### Excel Integration
- Appends **only new data**
- Uses a **checkpoint file** to prevent re-processing
- Automatically creates the Excel file and headers
- Safe for **daily scheduled execution**
- Clear command-line logging and success messages

---

## Repository Contents

| File | Description |
|-----|------------|
| `ZebraScanner.py` | Receives DS3678 batch uploads and appends validated serial numbers to text |
| `SerialToExcel.py` | Appends only new text data into Excel |
| `README.md` | Project documentation |

---

## Script 1 — ZebraScanner.py

### Purpose
Receives batch uploads from the DS3678 scanner and appends validated serial numbers to a text file.

### Behavior
- Waits for input to stop (batch complete)
- Extracts and validates serial numbers
- Appends to a persistent text file
- Clears input after processing

### Example Output (`psion.txt`) 

03/01/2026,19:05:44,123456789012345678
03/01/2026,19:05:44,987654321098765432


### Configuration
```python
OUTPUT_FILE = r"C:\Temp\psion.txt"
DELAY_MS = 2000
DELIM = ","

## Script 2 — SerialToExcel.py

### Purpose

`SerialToExcel.py` is a **scheduled ingestion script** that reads new scan records from the append-only text log and appends them into an Excel workbook.

It is designed to be run **daily via Windows Task Scheduler** and guarantees that **only new data** is added to Excel on each run.

---

### What This Script Does

- Reads scan records from the text log created by `ZebraScanner.py`
- Appends **only new records** into Excel
- Automatically creates the Excel file if it does not exist
- Creates headers **one time only**
- Skips invalid or malformed lines
- Prints status and success messages to the command window

---

### Input Format (Text File)

The script expects each line in the text file to follow this format:
MM/DD/YYYY,HH:MM:SS,Serial Number


Example:
03/01/2026,19:05:44,123456789012345678
03/01/2026,19:05:44,987654321098765432


---

### Excel Output

The Excel file contains the following headers:

| Date | Time | Serial Number |
|------|------|---------------|

Each new row appended corresponds to one valid serial number from the text file.

---

### Duplicate Prevention Strategy

This script **does not scan Excel for duplicates**.

Instead, it uses a **checkpoint file** that stores the last byte offset read from the text file.

#### How it works:
1. On startup, the script loads the last saved byte offset
2. It reads the text file starting from that position
3. Only newly appended lines are processed
4. After a successful Excel save, the checkpoint is updated

This method is:
- Fast
- Deterministic
- Safe for large files
- Immune to Excel formatting changes

---

### Configuration

Paths and settings can be configured at the top of the script:

```python
TEXT_FILE = r"C:\Temp\psion.txt"
EXCEL_FILE = r"C:\Temp\psion_log.xlsx"
CHECKPOINT_FILE = r"C:\Temp\psion_checkpoint.json"

SHEET_NAME = "Log"
HEADERS = ["Date", "Time", "Serial Number"]
DELIM = ","

### Command-Line Output

When run manually or via **Windows Task Scheduler**, the script prints clear status messages to the command window, including:

- Number of new lines read
- Number of valid rows appended
- Any invalid lines skipped
- Final success confirmation

#### Example Output
New lines read: 120
Valid rows appended: 120
Appended rows 501 to 620 in sheet 'Log'
SUCCESS: Daily append completed.


---

### Scheduling (Recommended)

Use **Windows Task Scheduler** to run this script on a daily basis.

#### Suggested Configuration
- **Trigger:** Daily
- **Action:**
python.exe SerialToExcel.py
- **Start in:**
path\to\script\directory

No scheduling logic is embedded in the script itself.

---

### Design Notes

- Excel headers are written **once only**
- Existing data is **never overwritten**
- The checkpoint is updated **only after a successful Excel save**
- Safe to rerun multiple times per day
- Performs reliably even as the text file grows large

---

### Common Use Cases

- Daily production traceability logs
- QA or audit reporting
- Offline scanning workflows
- Pre-MES data staging
