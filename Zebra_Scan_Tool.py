import tkinter as tk
import os
import re
from datetime import datetime

# ================= CONFIG =================
OUTPUT_FILE = r"F:\LFT\98_FCGF_Shares\17_Supply_Chain\End_of_Line\psion.txt"
DELAY_MS = 2000        # wait 2 seconds after last input (good for batch dumps)
DELIM = ","            # comma-delimited text file
# ==========================================

# -------- Window setup --------
root = tk.Tk()
root.geometry("1200x300")
root.title("DS3678 Batch Receiver")
root.attributes("-topmost", True)
root.resizable(False, False)

tk.Label(
    root,
    text="Awaiting Serial Numbers....",
    font=("Courier", 14)
).pack(pady=10)

entry = tk.Entry(root, width=120)
entry.pack(padx=10)
entry.focus_set()

status = tk.StringVar(value="Waiting for batch upload... \nPLEASE CLICK ANYWHERE ON THIS WINDOW BEFORE RETURNING SCANNER TO DOCK!!!")
tk.Label(
    root,
    textvariable=status,
    font=("Courier", 16)
).pack(pady=8)

after_id = None

# -------- Force focus & topmost --------
def force_focus():
    try:
        root.attributes("-topmost", True)
        entry.focus_force()
        entry.icursor(tk.END)
    except tk.TclError:
        pass
    root.after(500, force_focus)

def on_focus_out(event):
    entry.focus_force()
    entry.icursor(tk.END)

root.bind("<FocusOut>", on_focus_out)
entry.bind("<FocusOut>", on_focus_out)

# -------- Extract valid SERNRs --------
def extract_sernrs(raw: str) -> list[str]:
    digit_runs = re.findall(r"\d+", raw)

    results = []
    seen = set()

    for d in digit_runs:
        # Must be at least one full SERNR
        if len(d) < 18:
            continue

        # Length must be divisible by 18
        if len(d) % 18 != 0:
            continue

        # Split into 18-digit chunks
        for i in range(0, len(d), 18):
            sernr = d[i:i + 18]
            if sernr not in seen:
                seen.add(sernr)
                results.append(sernr)

    return results

# -------- Process batch --------
def process_entry():
    global after_id

    raw = entry.get()
    sernrs = extract_sernrs(raw)

    if sernrs:
        os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)

        now = datetime.now()
        date_str = now.strftime("%m/%d/%Y")
        time_str = now.strftime("%H:%M:%S")

        with open(OUTPUT_FILE, "a", encoding="utf-8") as f:
            for s in sernrs:
                f.write(f"{date_str}{DELIM}{time_str}{DELIM}{s}\n")

        status.set(f"Appended {len(sernrs)} Serial Number(s)")
    else:
        status.set("No valid 18-digit Serial Number found")

    entry.delete(0, tk.END)
    entry.focus_force()
    after_id = None

# -------- Restart timer on every input --------
def schedule_processing(event=None):
    global after_id
    if after_id is not None:
        root.after_cancel(after_id)
    after_id = root.after(DELAY_MS, process_entry)

entry.bind("<KeyRelease>", schedule_processing)

# -------- Start focus enforcement --------
force_focus()

root.mainloop()