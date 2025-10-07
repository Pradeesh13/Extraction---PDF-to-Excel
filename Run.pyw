import customtkinter as ctk
from tkinter import filedialog, messagebox
import configparser
import subprocess
import os, sys

# ------------ Constants ------------
PROJECT_TITLE = "PDF TO EXCEL CONVERTER"
PYTHON_SCRIPTS = [
    "Pdf_to_Excel.py",
    "Excel_to_ini_new.py",
    "CellValueInserting.py"
]

PRIMARY = "#223447"
PRIMARY_HOVER = "#1B2B3C"
SURFACE = "#F6F8FA"
FIELD_BG = "#2B3E52"
FIELD_TXT = "#EAF1F7"
LABEL_CLR = "#223447"
PLACEHOLDER = "#B7C4CF"
TEXT_SIZE = 14

# ------------ Resource Path Helper ------------
def resource_path(relative_path):
    """Get absolute path to resource (works for dev and PyInstaller exe)."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

PATH_INI_FILE = resource_path(os.path.join("Info", "Config", "path.ini"))

# ------------ UI Setup ------------
root = ctk.CTk()
root.title(PROJECT_TITLE)
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("dark-blue")
root.configure(fg_color=SURFACE)

# Center window
window_width, window_height = 640, 430
x_pos = (root.winfo_screenwidth() // 2) - (window_width // 2)
y_pos = (root.winfo_screenheight() // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")
root.resizable(False, False)

# ------------ Title ------------
headingFrame = ctk.CTkFrame(master=root, fg_color=PRIMARY, corner_radius=0)
headingFrame.pack(fill="x")
ctk.CTkLabel(
    master=headingFrame,
    text=PROJECT_TITLE,
    text_color="white",
    font=("Segoe UI", 24, "bold")
).pack(padx=24, pady=16)

# ------------ Content Frame ------------
contentFrame = ctk.CTkFrame(master=root, fg_color=SURFACE, corner_radius=0)
contentFrame.pack(fill="both", padx=24, pady=24)

# Input PDF
ctk.CTkLabel(contentFrame, text="Choose Input PDF File",
             text_color=LABEL_CLR, font=("Segoe UI", TEXT_SIZE, "bold")
).pack(padx=15, pady=(6, 0), anchor="w")

inputEntry = ctk.CTkEntry(contentFrame,
                          placeholder_text="Select PDF file...",
                          width=560, text_color=FIELD_TXT, fg_color=FIELD_BG,
                          font=("Segoe UI", TEXT_SIZE))
inputEntry.pack(padx=0, pady=(0, 6))

def browse_input():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        inputEntry.delete(0, "end")
        inputEntry.insert(0, file_path)

ctk.CTkButton(contentFrame, text="Browse", command=browse_input,
              fg_color=PRIMARY, hover_color=PRIMARY_HOVER, text_color="white",
              width=110, height=25).pack(padx=15, pady=(0, 14), anchor="e")

# Output Excel
ctk.CTkLabel(contentFrame, text="Choose Output Excel File",
             text_color=LABEL_CLR, font=("Segoe UI", TEXT_SIZE, "bold")
).pack(padx=15, pady=(6, 0), anchor="w")

outputEntry = ctk.CTkEntry(contentFrame,
                           placeholder_text="Save Excel file as...",
                           width=560, text_color=FIELD_TXT, fg_color=FIELD_BG,
                           font=("Segoe UI", TEXT_SIZE))
outputEntry.pack(padx=0, pady=(0, 6))

def browse_output():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        outputEntry.delete(0, "end")
        outputEntry.insert(0, file_path)

ctk.CTkButton(contentFrame, text="Browse", command=browse_output,
              fg_color=PRIMARY, hover_color=PRIMARY_HOVER, text_color="white",
              width=110, height=25).pack(padx=15, pady=(0, 14), anchor="e")

# ------------ Run Scripts Logic ------------
def run_scripts():
    input_path = inputEntry.get().strip()
    output_path = outputEntry.get().strip()

    if not input_path or not output_path:
        messagebox.showerror("Error", "Please select both input PDF and output Excel file.")
        return

    config = configparser.ConfigParser()
    if not config.has_section('input'):
        config.add_section('input')
    if not config.has_section('output'):
        config.add_section('output')
    config.set('input', 'path', input_path)
    config.set('output', 'path', output_path)

    os.makedirs(os.path.dirname(PATH_INI_FILE), exist_ok=True)
    with open(PATH_INI_FILE, 'w') as f:
        config.write(f)

    try:
        for script in PYTHON_SCRIPTS:
            script_path = resource_path(script)
            result = subprocess.run([sys.executable, script_path], check=True, capture_output=True, text=True)
            print(f"Output from {script}:\n{result.stdout}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Script Error", f"Error running {script}:\n{e.stderr}")
        return

    # Cleanup
    config.set('input', 'path', "")
    config.set('output', 'path', "")
    with open(PATH_INI_FILE, 'w') as f:
        config.write(f)

    inputEntry.delete(0, "end")
    outputEntry.delete(0, "end")

    messagebox.showinfo("Success", "Conversion completed successfully!")

# ------------ Convert Button ------------
ctk.CTkButton(contentFrame, text="CONVERT", command=run_scripts,
              fg_color=PRIMARY, hover_color=PRIMARY_HOVER, text_color="white",
              width=320, height=35, font=("Segoe UI", 15, "bold")
).pack(padx=0, pady=20, anchor="n")

root.mainloop()
