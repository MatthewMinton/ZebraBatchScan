# Zebra Dock Scan Tool, DS3678 Batch Receiver, TXT ➜ Excel Logger

A lightweight, production-friendly traceability utility that captures **18-digit SERNR serial numbers** from a Zebra **DS3678** dock, writes them to a **comma-delimited TXT log**, then appends new scans into a **central Excel log** with **duplicate protection** and a **checkpoint** so it is safe to run repeatedly.

This repo includes:

- **Tkinter “always-on-top” receiver UI** for batch uploads
- **TXT log writer** (date, time, serial)
- **Excel daily append utility** with:
  - checkpointing (resume where you left off)
  - duplicate serial prevention (based on Excel contents)
  - invalid line filtering

---

## Table of Contents

- [Project Overview](#project-overview)
- [How It Works](#how-it-works)
- [Files Included](#files-included)
- [Data Formats](#data-formats)
- [Requirements](#requirements)
- [Setup](#setup)
- [Usage](#usage)
- [Scheduling the Excel Append Script](#scheduling-the-excel-append-script)
- [Duplicate Handling](#duplicate-handling)
- [Troubleshooting](#troubleshooting)
- [Design Notes](#design-notes)
- [Future Improvements](#future-improvements)
- [Author](#author)
- [License](#license)

---

## Project Overview

This toolchain supports an **offline scan workflow** where:

1. A Zebra DS3678 is docked, then uploads a **batch** of scanned serial numbers.
2. A small Tkinter window receives the pasted/batched data and extracts valid **18-digit numeric serials**.
3. Each extracted serial is appended to a centralized text log file:
   - `MM/DD/YYYY,HH:MM:SS,18-digit-serial`
4. A second script reads *only new lines* from that text file (using a checkpoint) and appends them into a shared Excel workbook.

---

## How It Works

### 1) Batch Receiver (Tkinter UI)

- Window is forced **topmost** and the cursor focus is kept in the entry box.
- Raw batch text is parsed for digit runs.
- Only digit runs that can be split cleanly into **18-digit chunks** are accepted.
- Each valid serial is written to the TXT log with the current date and time.

### 2) TXT ➜ Excel Append Script

- Reads `psion_checkpoint.json` to determine the last-read byte offset.
- Seeks directly to that offset in the TXT file and reads *only newly added lines*.
- Ensures the Excel workbook exists, creates it if not.
- Loads existing serial numbers from Excel column 3 (Serial Number) into a set.
- Appends only **new unique serials**.
- Updates checkpoint after successful processing.

---

## Files Included

### 1) DS3678 Batch Receiver (Tkinter)
A GUI tool that receives the dock upload and writes valid serials to the TXT log.

**Key behaviors**
- Always on top (`root.attributes("-topmost", True)`)
- Re-focuses input automatically
- Extracts serials from pasted/batched content
- Writes `Date,Time,Serial` to TXT

### 2) TXT ➜ Excel Daily Append Script
Reads the TXT log and appends new data to Excel, safely and repeatably.

**Key behaviors**
- Checkpoint-based reading (byte offset + file modified time)
- Duplicate protection based on Excel serial column
- Invalid/empty line filtering
- Creates Excel file + headers if missing

### 3) `psion_checkpoint.json`
Stores how far the append script has processed in the TXT log.

Example:
```json
{
  "offset": 760,
  "text_mtime": 1772479749.420252,
  "updated": "03/02/2026 14:29:21"
}


