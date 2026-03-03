![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-blue)
![GUI](https://img.shields.io/badge/GUI-Tkinter-lightgrey)
![Excel](https://img.shields.io/badge/Excel-openpyxl-green)
![Status](https://img.shields.io/badge/Status-Production--Ready-brightgreen)

![Traceability](https://img.shields.io/badge/Use--Case-Manufacturing%20Traceability-orange)
![Barcode](https://img.shields.io/badge/Scanner-Zebra%20DS3678-black)
![Offline](https://img.shields.io/badge/Workflow-Offline%20Batch-lightblue)

![Duplicate Safe](https://img.shields.io/badge/Duplicates-Prevented-success)
![Checkpointing](https://img.shields.io/badge/Checkpointing-Enabled-informational)
![UTF--8](https://img.shields.io/badge/Encoding-UTF--8-blueviolet)

# Zebra Dock Scan Tool, DS3678 Batch Receiver, TXT ➜ Excel Logger

A lightweight, production-friendly traceability utility that captures **18-digit SERNR serial numbers** from a Zebra **DS3678** dock, writes them to a **comma-delimited TXT log**, then appends new scans into a **central Excel log** with **duplicate protection** and a **checkpoint**, making it safe to run repeatedly.

---

## This Repository Includes

- Tkinter **always-on-top receiver UI** for batch uploads  
- **TXT log writer** (date, time, serial)  
- **Excel daily append utility** with:
  - checkpointing (resume where you left off)
  - duplicate serial prevention (based on Excel contents)
  - invalid line filtering

---

## Table of Contents

- Project Overview  
- How It Works  
- Files Included  
- Data Formats  
- Requirements  
- Setup  
- Usage  
- Scheduling the Excel Append Script  
- Duplicate Handling  
- Troubleshooting  
- Design Notes  
- Future Improvements  
- Author  
- License  

---

## Project Overview

This toolchain supports an **offline scan workflow** where:

- A Zebra DS3678 is docked, then uploads a **batch** of scanned serial numbers
- A small Tkinter window receives the pasted or batched data and extracts valid **18-digit numeric serials**
- Each extracted serial is appended to a centralized text log file using the format:

MM/DD/YYYY,HH:MM:SS,18-digit-serial


- A second script reads **only new lines** from that text file (using a checkpoint) and appends them into a shared Excel workbook

---

## How It Works

### 1) Batch Receiver (Tkinter UI)

- The window is forced **always on top**, and cursor focus is kept in the entry box
- Raw batch text is parsed for digit runs
- Only digit runs that can be split cleanly into **18-digit chunks** are accepted
- Each valid serial is written to the TXT log with the current date and time

---

### 2) TXT ➜ Excel Append Script

- Reads `psion_checkpoint.json` to determine the last-read byte offset
- Seeks directly to that offset in the TXT file and reads **only newly added lines**
- Ensures the Excel workbook exists, creating it if necessary
- Loads existing serial numbers from **Excel column 3 (Serial Number)** into a set
- Appends **only new unique serials**
- Updates the checkpoint after successful processing

---

## Files Included

### 1) DS3678 Batch Receiver (Tkinter)

A GUI tool that receives the dock upload and writes valid serials to the TXT log.

**Key behaviors**
- Always on top (`root.attributes("-topmost", True)`)
- Automatically re-focuses input
- Extracts serials from pasted or batched content
- Writes `Date,Time,Serial` to the TXT log

---

### 2) TXT ➜ Excel Daily Append Script

Reads the TXT log and appends new data to Excel in a safe and repeatable way.

**Key behaviors**
- Checkpoint-based reading (byte offset + file modified time)
- Duplicate protection based on Excel serial column
- Invalid and empty line filtering
- Creates the Excel file and headers if missing

---

### 3) psion_checkpoint.json

Stores how far the append script has processed in the TXT log.

**Example**
```json
{
  "offset": 760,
  "text_mtime": 1772479749.420252,
  "updated": "03/02/2026 14:29:21"
}
```
## Data Formats

### TXT Log Line Format

Each line in the log must match the following format:
MM/DD/YYYY,HH:MM:SS,18-digit-serial


**Example**
03/02/2026,14:29:21,123456789012345678


---

### Excel Columns

The Excel workbook sheet **Log** uses the following columns:

| Date | Time | Serial Number |
|------|------|---------------|

---

## Requirements

- **Windows**  
  Paths are configured for Windows drive letters and network shares

- **Python 3.x**

### Required Packages

- `openpyxl`

Install with:

```bash
pip install openpyxl
```
tkinter ships with most Python installs on Windows.
If it is missing, install the full Python distribution or enable the Tk components.

## Setup

### 1) Configure Paths

Both scripts rely on **hard-coded file paths**.  
Update these values to match your environment.

#### Receiver Script

- `OUTPUT_FILE`

#### Append Script

- `TEXT_FILE`
- `EXCEL_FILE`
- `CHECKPOINT_FILE`

---

## Usage

### Step 1: Run the Batch Receiver

Run the Tkinter receiver on the workstation connected to the dock:

```bash
python Zebra_Scan_Tool.py
```

### Console Output Includes
- Number of new lines read
- Valid rows appended
- Duplicate serial nummbers skipped
- Invalid lines skipped
- Appended row range

---

## Scheduling the Excel Append Script

**Recommended:** Use **Windows Task Scheduler** to run the script on a cadence (daily, hourly, etc.).

### Suggested Configuration

### Trigger

- Daily, or every X minutes

### Action
- **Program/script:** python.exe
- **Add arguments:** Zebra_To_Excel.py
- **Start in:** directory containing the script

No scheduling logic is embedded in the script.

It is intentionally designed to be scheduled externally.

---

## Duplicate Handling

Duplicates are prevented at the **Excel level.**
- Before appending, the script loafs all existing serals from **column 3.**
- Any serial already present is skipped
- serial accepted during the current run are added immediately, preventing duplicates within the same run.

This makes the append script **safe to run repeatedly** without re-adding serial numbers.

---

## Troubleshooting

### Receiver Window is not Capturing Scans

- Click inside the receiver window at least once before docking the scanner upload
- Some environments require initial focus action
- The UI enforces focus, but Windows policies may still allow other applications to steal focus

---

### No Data to Append to Excel

- Confirm psion.txt exists and is being written to
- confirm checkpint file is not pointing past the end of the txt file
- If the txt file was truncated or rewritten, the script automatically resets the checkpoint

---

### Duplcates Are Still Showing in Excel

This should not happen unless:
- serials are not stored consistently as 18-digit strings
- serials were manually edited or entered with spaces, formatting changes, or leading zeros removed

---

### Design Notes

- The chekpoint is updated **only after a successful Excel save.**
- Excel headers are written once, and the sheet is created if it does not exist
- Invalid lines are ignored instead of crashing the job
- Files are handled using **UTF-8 encoding**

---

### Author

**Matthew Minton**

Controls Engineer, Manufacturing Engineering

March 2, 2026

---

### License

Internal use, customize as needed for your environment.

