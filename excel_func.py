import tkinter as tk
from tkinter import filedialog, simpledialog, ttk
import pandas as pd

def get_excel_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
    )
    root.destroy()
    return file_path

def get_pivot_params():
    root = tk.Tk()
    root.withdraw()
    index = simpledialog.askstring("Input", "Enter column names for index (comma separated):")
    columns = simpledialog.askstring("Input", "Enter column names for columns (comma separated):")
    values = simpledialog.askstring("Input", "Enter column names for values (comma separated):")
    aggfunc = simpledialog.askstring("Input", "Enter aggregation function (e.g., sum, mean, count):")
    fill_value = simpledialog.askstring("Input", "Enter fill value (leave blank for None):")
    root.destroy()
    # Process inputs
    index = [i.strip() for i in index.split(",")] if index else None
    columns = [c.strip() for c in columns.split(",")] if columns else None
    values = [v.strip() for v in values.split(",")] if values else None
    aggfunc = aggfunc.strip() if aggfunc else 'sum'
    fill_value = fill_value.strip() if fill_value else None
    return index, columns, values, aggfunc, fill_value

def analyze_columns(file_path):
    df = pd.read_excel(file_path)
    return df

def display_pivot_table(pivot_df):
    root = tk.Tk()
    root.title("Pivot Table")
    tree = ttk.Treeview(root)
    tree.pack(expand=True, fill='both')
    # Flatten MultiIndex columns if present
    cols = [" ".join([str(i) for i in col]).strip() if isinstance(col, tuple) else str(col) for col in pivot_df.columns]
    tree['columns'] = cols
    tree['show'] = 'headings'
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, anchor='center')
    for idx, row in pivot_df.iterrows():
        tree.insert('', 'end', values=list(row))
    root.mainloop()

if __name__ == "__main__":
    selected_file = get_excel_file()
    if selected_file:
        df = analyze_columns(selected_file)
        index, columns, values, aggfunc, fill_value = get_pivot_params()
        # Convert fill_value to appropriate type
        if fill_value == '':
            fill_value = None
        else:
            try:
                fill_value = float(fill_value)
            except ValueError:
                pass
        pivot_table = pd.pivot_table(
            df,
            index=index,
            columns=columns,
            values=values,
            aggfunc=aggfunc,
            fill_value=fill_value
        )
        display_pivot_table(pivot_table)
    else:
        print("No file selected")
