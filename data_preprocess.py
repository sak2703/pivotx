import tkinter as tk
from tkinter import filedialog
import pandas as pd

def get_excel_file():
    """Open file dialog to select Excel file"""
    root = tk.Tk()
    root.withdraw()  # Hide main window
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
    )
    return file_path

def analyze_columns(file_path):
    """Read Excel file and return column data types"""
    df = pd.read_excel(file_path)
    return df.dtypes

if __name__ == "__main__":
    #print("lol")
    selected_file = get_excel_file()
    
    if selected_file:
        print(f"\nSelected file: {selected_file}")
        column_types = analyze_columns(selected_file)
        print("\nColumn Data Types:")
        print(column_types)
    else:
        print("No file selected")
        