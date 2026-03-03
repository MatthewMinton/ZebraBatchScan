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

---

## Data Format
### TXT Log Line Format
Each line in the log must match the following format:
MM/DD/YYYY,HH:MM:SS,18-digit-serial
