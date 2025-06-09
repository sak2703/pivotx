import tkinter as tk
from tkinter import filedialog, ttk, messagebox
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

def get_pivot_params(df):
    # Helper to get columns from DataFrame
    columns = list(df.columns)
    selected = {'index': [], 'columns': [], 'values': [], 'aggfunc': 'sum', 'fill_value': None}
    done = {'status': False}

    def on_submit():
        # Gather selections
        selected['index'] = [col for col, var in index_vars.items() if var.get()]
        selected['columns'] = [col for col, var in columns_vars.items() if var.get()]
        selected['values'] = [col for col, var in values_vars.items() if var.get()]
        selected['aggfunc'] = aggfunc_var.get()
        fv = fill_value_var.get()
        selected['fill_value'] = None if fv.strip() == '' else fv
        done['status'] = True
        param_win.destroy()

    def on_cancel():
        done['status'] = False
        param_win.destroy()

    param_win = tk.Tk()
    param_win.title("Select Pivot Table Parameters")

    # Index selection
    tk.Label(param_win, text="Select index columns:").grid(row=0, column=0, sticky='w')
    index_vars = {col: tk.BooleanVar() for col in columns}
    for i, col in enumerate(columns):
        tk.Checkbutton(param_win, text=col, variable=index_vars[col]).grid(row=1, column=i, sticky='w')

    # Columns selection
    tk.Label(param_win, text="Select columns:").grid(row=2, column=0, sticky='w')
    columns_vars = {col: tk.BooleanVar() for col in columns}
    for i, col in enumerate(columns):
        tk.Checkbutton(param_win, text=col, variable=columns_vars[col]).grid(row=3, column=i, sticky='w')

    # Values selection
    tk.Label(param_win, text="Select values:").grid(row=4, column=0, sticky='w')
    values_vars = {col: tk.BooleanVar() for col in columns}
    for i, col in enumerate(columns):
        tk.Checkbutton(param_win, text=col, variable=values_vars[col]).grid(row=5, column=i, sticky='w')

    # Aggregation function
    tk.Label(param_win, text="Aggregation function:").grid(row=6, column=0, sticky='w')
    aggfunc_var = tk.StringVar(value='sum')
    aggfunc_options = ['sum', 'mean', 'count', 'min', 'max', 'median', 'std', 'var']
    aggfunc_menu = ttk.Combobox(param_win, textvariable=aggfunc_var, values=aggfunc_options, state='readonly')
    aggfunc_menu.grid(row=6, column=1, sticky='w')

    # Fill value
    tk.Label(param_win, text="Fill value (optional):").grid(row=7, column=0, sticky='w')
    fill_value_var = tk.StringVar()
    tk.Entry(param_win, textvariable=fill_value_var).grid(row=7, column=1, sticky='w')

    # Buttons
    tk.Button(param_win, text="Submit", command=on_submit).grid(row=8, column=0, pady=10)
    tk.Button(param_win, text="Cancel", command=on_cancel).grid(row=8, column=1, pady=10)

    param_win.mainloop()

    if not done['status']:
        return None, None, None, None, None

    # Convert empty lists to None for pandas
    for key in ['index', 'columns', 'values']:
        if not selected[key]:
            selected[key] = None
    # Try to convert fill_value to numeric if possible
    fv = selected['fill_value']
    if fv is not None:
        try:
            selected['fill_value'] = float(fv)
        except ValueError:
            pass
    return selected['index'], selected['columns'], selected['values'], selected['aggfunc'], selected['fill_value']

def display_pivot_table(pivot_df):
    root = tk.Tk()
    root.title("Pivot Table")
    tree = ttk.Treeview(root)
    tree.pack(expand=True, fill='both')
    # Flatten MultiIndex columns if present
    if isinstance(pivot_df.columns, pd.MultiIndex):
        cols = [" ".join([str(i) for i in col]).strip() for col in pivot_df.columns]
    else:
        cols = [str(col) for col in pivot_df.columns]
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
    if not selected_file:
        print("No file selected")
        exit()
    df = pd.read_excel(selected_file)
    params = get_pivot_params(df)
    if any(param is None for param in params):
        print("Pivot table creation cancelled by user.")
        exit()
    index, columns, values, aggfunc, fill_value = params
    try:
        pivot_table = pd.pivot_table(
            df,
            index=index,
            columns=columns,
            values=values,
            aggfunc=aggfunc,
            fill_value=fill_value
        )
        display_pivot_table(pivot_table)
    except Exception as e:
        messagebox.showerror("Error", f"Could not create pivot table:\n{e}")
