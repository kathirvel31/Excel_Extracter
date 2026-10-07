import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

from excel_reader import get_sheet_names, load_sheet_data
from field_search import filter_fields
from data_processor import process_data
from export import export_to_excel

app_state = {
    "file_paths": [],
    "df_original": None,
    "df_filtered": None,
    "all_columns": [],
    "selected_columns": set()
}

def upload_file():
    """
    Handles the file upload process. Opens a file dialog for the user to select one or multiple Excel files.
    Updates the application state and UI accordingly.
    """
    filepaths = filedialog.askopenfilenames(
        title="Select Excel File(s)",
        filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
    )
    if not filepaths:
        return
    
    app_state["file_paths"] = list(filepaths)
    
    if len(filepaths) == 1:
        lbl_filename.config(text=filepaths[0].split("/")[-1])
        combo_sheets.config(state="readonly")
        try:
            sheets = get_sheet_names(filepaths[0])
            combo_sheets['values'] = sheets
            if sheets:
                combo_sheets.current(0)
                on_sheet_select(None)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file:\n{str(e)}")
    else:
        lbl_filename.config(text=f"{len(filepaths)} files selected")
        combo_sheets.set("Multiple Files (Sheet 1)")
        combo_sheets.config(state="disabled")
        load_multiple_files(filepaths)

def load_multiple_files(filepaths):
    """
    Loads data from multiple Excel files and combines them into a single DataFrame.
    Updates the application state with the combined data and all unique columns.
    """
    dfs = []
    for fp in filepaths:
        try:
            df = load_sheet_data(fp, 0)
            dfs.append(df)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file {fp.split('/')[-1]}:\n{str(e)}")
            return
            
    if dfs:
        combined_df = pd.concat(dfs, ignore_index=True)
        app_state["df_original"] = combined_df
        
        all_cols = []
        for df in dfs:
            for col in df.columns:
                if col not in all_cols:
                    all_cols.append(col)
                    
        app_state["all_columns"] = all_cols
        app_state["selected_columns"].clear()
        
        try:
            entry_search_var.set("")
        except NameError:
            pass
            
        update_preview_table(combined_df)
        update_field_listbox(app_state["all_columns"])

def on_sheet_select(event):
    """
    Event handler for sheet selection from the combobox.
    Loads the data from the selected sheet and updates the preview table and field listbox.
    """
    sheet = combo_sheets.get()
    if not sheet or not app_state.get("file_paths"):
        return
        
    try:
        df = load_sheet_data(app_state["file_paths"][0], sheet)
        app_state["df_original"] = df
        app_state["all_columns"] = list(df.columns)
        app_state["selected_columns"].clear()
        
        try:
            entry_search_var.set("")
        except NameError:
            pass
            
        update_preview_table(df)
        update_field_listbox(app_state["all_columns"])
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load sheet:\n{str(e)}")

def update_preview_table(df):
    """
    Updates the preview treeview with the first 200 rows of the provided DataFrame.
    """
    for row in tree_preview.get_children():
        tree_preview.delete(row)
        
    tree_preview["column"] = list(df.columns)
    tree_preview["show"] = "headings"
    
    for col in tree_preview["columns"]:
        tree_preview.heading(col, text=col)
        tree_preview.column(col, width=120)
        
    # show top 200 rows for preview to keep it fast
    df_preview = df.head(200)
    for _, row in df_preview.iterrows():
        # Safely convert to string, skipping true NaNs to show as blank
        tree_preview.insert("", "end", values=[str(x) if pd.notna(x) and str(x).lower() != "nan" else "" for x in row])

def on_search(*args):
    """
    Event handler for the search entry. Filters the available fields listbox based on the search text.
    """
    search_text = entry_search_var.get()
    filtered = filter_fields(app_state["all_columns"], search_text)
    update_field_listbox(filtered)

def update_field_listbox(fields):
    """
    Updates the listbox with the provided list of fields and restores any previous selections.
    """
    listbox_fields.delete(0, tk.END)
    for i, field in enumerate(fields):
        listbox_fields.insert(tk.END, field)
        if field in app_state["selected_columns"]:
            listbox_fields.selection_set(i)

def on_listbox_select(event):
    """
    Event handler for selection changes in the field listbox.
    Updates the application state with the currently selected columns.
    """
    selected_indices = listbox_fields.curselection()
    selected_in_view = set([listbox_fields.get(i) for i in selected_indices])
    displayed = set(listbox_fields.get(0, tk.END))
    
    for item in selected_in_view:
        app_state["selected_columns"].add(item)
        
    for item in displayed:
        if item not in selected_in_view and item in app_state["selected_columns"]:
            app_state["selected_columns"].remove(item)

def fetch_data():
    """
    Processes the loaded data by extracting selected columns, filtering empty rows,
    and sorting based on user preferences. Updates the result table and switches to the result tab.
    """
    if app_state["df_original"] is None:
        messagebox.showwarning("Warning", "Please upload an Excel file first.")
        return
        
    # To maintain the order of selection as they appear in the original file
    selected_cols = [col for col in app_state["all_columns"] if col in app_state["selected_columns"]]
    
    if not selected_cols:
        messagebox.showwarning("Warning", "Please select at least one field.")
        return
        
    sort_order = sort_var.get()
    
    try:
        df_filtered = process_data(app_state["df_original"], selected_cols, sort_order)
        app_state["df_filtered"] = df_filtered
        update_result_table(df_filtered)
        notebook.select(tab_result)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process data:\n{str(e)}")

def on_row_search(*args):
    """
    Event handler for the row search entry in the result tab.
    Filters the displayed rows in the result treeview based on the search text.
    """
    try:
        search_text = entry_row_search_var.get().lower()
    except NameError:
        return
        
    df = app_state.get("df_filtered")
    if df is None:
        return
        
    if not search_text:
        _update_result_tree(df)
        return
        
    # Search for rows where any column starts with the search text (case-insensitive)
    mask = df.astype(str).apply(lambda col: col.str.lower().str.startswith(search_text, na=False)).any(axis=1)
    df_searched = df[mask]
    _update_result_tree(df_searched)

def update_result_table(df):
    """
    Wrapper function to update the result treeview. Handles clearing the search box if necessary.
    """
    try:
        if entry_row_search_var.get() != "":
            entry_row_search_var.set("") # This triggers the trace to update the tree
            return
    except NameError:
        pass
    _update_result_tree(df)

def _update_result_tree(df):
    """
    Internal function to populate the result treeview with all rows from the provided DataFrame.
    """
    for row in tree_result.get_children():
        tree_result.delete(row)
        
    tree_result["column"] = list(df.columns)
    tree_result["show"] = "headings"
    
    for col in tree_result["columns"]:
        tree_result.heading(col, text=col)
        tree_result.column(col, width=150)
        
    for _, row in df.iterrows():
        tree_result.insert("", "end", values=[str(x) if pd.notna(x) and str(x).lower() != "nan" else "" for x in row])

def export_file():
    """
    Handles the export process. Opens a file dialog for the user to select a save location
    and exports the filtered DataFrame to an Excel file.
    """
    if app_state["df_filtered"] is None:
        messagebox.showwarning("Warning", "No filtered data to export. Please fetch data first.")
        return
        
    filepath = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile="filtered_result.xlsx",
        title="Save Exported Data",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    if not filepath:
        return
        
    try:
        export_to_excel(app_state["df_filtered"], filepath)
        messagebox.showinfo("Success", "Data exported successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export data:\n{str(e)}")

def build_gui():
    """
    Constructs and configures the main Tkinter GUI application, including all widgets and layout.
    Starts the Tkinter main event loop.
    """
    global lbl_filename, combo_sheets, entry_search_var, entry_search, listbox_fields, sort_var
    global tree_preview, tree_result, notebook, tab_result
    global entry_row_search_var
    
    root = tk.Tk()
    root.title("Excel Field Extractor & Sorter")
    root.geometry("950x650")
    
    # Left Panel
    frame_left = tk.Frame(root, width=280, bg="#e8ecef", padx=15, pady=15)
    frame_left.pack(side=tk.LEFT, fill=tk.Y)
    
    # Right Panel
    frame_right = tk.Frame(root, padx=10, pady=10)
    frame_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
    
    lbl_title = tk.Label(frame_left, text="Excel Data Fetcher", font=("Segoe UI", 16, "bold"), bg="#e8ecef", fg="#333")
    lbl_title.pack(pady=(0, 15))
    
    btn_upload = tk.Button(frame_left, text="Upload Excel File", command=upload_file, bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold"), relief=tk.FLAT, pady=5)
    btn_upload.pack(fill=tk.X, pady=5)
    
    lbl_filename = tk.Label(frame_left, text="No file selected", bg="#e8ecef", fg="#555", font=("Segoe UI", 9, "italic"))
    lbl_filename.pack(pady=5)
    
    tk.Label(frame_left, text="Select Sheet:", bg="#e8ecef", font=("Segoe UI", 10)).pack(anchor=tk.W, pady=(15, 2))
    combo_sheets = ttk.Combobox(frame_left, state="readonly", font=("Segoe UI", 9))
    combo_sheets.pack(fill=tk.X, pady=2)
    combo_sheets.bind("<<ComboboxSelected>>", on_sheet_select)
    
    tk.Label(frame_left, text="Search Field:", bg="#e8ecef", font=("Segoe UI", 10)).pack(anchor=tk.W, pady=(15, 2))
    entry_search_var = tk.StringVar()
    entry_search_var.trace_add("write", on_search)
    entry_search = tk.Entry(frame_left, textvariable=entry_search_var, font=("Segoe UI", 10))
    entry_search.pack(fill=tk.X, pady=2)
    
    tk.Label(frame_left, text="Available Fields (Select multiple):", bg="#e8ecef", font=("Segoe UI", 10)).pack(anchor=tk.W, pady=(15, 2))
    
    frame_listbox = tk.Frame(frame_left)
    frame_listbox.pack(fill=tk.BOTH, expand=True, pady=2)
    
    scrollbar = tk.Scrollbar(frame_listbox)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    listbox_fields = tk.Listbox(frame_listbox, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set, font=("Segoe UI", 10), selectbackground="#0078D7")
    listbox_fields.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    listbox_fields.bind("<<ListboxSelect>>", on_listbox_select)
    scrollbar.config(command=listbox_fields.yview)
    
    tk.Label(frame_left, text="Sort:", bg="#e8ecef", font=("Segoe UI", 10)).pack(anchor=tk.W, pady=(15, 2))
    sort_var = tk.StringVar(value="A-Z")
    frame_sort = tk.Frame(frame_left, bg="#e8ecef")
    frame_sort.pack(fill=tk.X, pady=2)
    tk.Radiobutton(frame_sort, text="A → Z", variable=sort_var, value="A-Z", bg="#e8ecef", font=("Segoe UI", 10)).pack(side=tk.LEFT)
    tk.Radiobutton(frame_sort, text="Z → A", variable=sort_var, value="Z-A", bg="#e8ecef", font=("Segoe UI", 10)).pack(side=tk.LEFT)
    
    btn_fetch = tk.Button(frame_left, text="Fetch Selected Fields", command=fetch_data, bg="#008CBA", fg="white", font=("Segoe UI", 10, "bold"), relief=tk.FLAT, pady=5)
    btn_fetch.pack(fill=tk.X, pady=(20, 5))
    
    btn_export = tk.Button(frame_left, text="Export to Excel", command=export_file, bg="#FF9800", fg="white", font=("Segoe UI", 10, "bold"), relief=tk.FLAT, pady=5)
    btn_export.pack(fill=tk.X, pady=5)
    
    # Setup Notebook (Tabs) for Right Panel
    style = ttk.Style()
    style.configure("TNotebook.Tab", font=("Segoe UI", 10))
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    style.configure("Treeview", font=("Segoe UI", 9))
    
    notebook = ttk.Notebook(frame_right)
    notebook.pack(fill=tk.BOTH, expand=True)
    
    tab_preview = tk.Frame(notebook)
    notebook.add(tab_preview, text=" Preview Original Data ")
    
    tab_result = tk.Frame(notebook)
    notebook.add(tab_result, text=" Filtered Result ")
    
    # Add Search Box for Rows in Result Tab
    frame_result_top = tk.Frame(tab_result, padx=5, pady=5)
    frame_result_top.pack(fill=tk.X)
    tk.Label(frame_result_top, text="Search Data (Alphabet/Text):", font=("Segoe UI", 10)).pack(side=tk.LEFT)
    entry_row_search_var = tk.StringVar()
    entry_row_search_var.trace_add("write", on_row_search)
    tk.Entry(frame_result_top, textvariable=entry_row_search_var, font=("Segoe UI", 10)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    
    # Preview Treeview
    scroll_prev_y = tk.Scrollbar(tab_preview)
    scroll_prev_y.pack(side=tk.RIGHT, fill=tk.Y)
    scroll_prev_x = tk.Scrollbar(tab_preview, orient=tk.HORIZONTAL)
    scroll_prev_x.pack(side=tk.BOTTOM, fill=tk.X)
    
    tree_preview = ttk.Treeview(tab_preview, yscrollcommand=scroll_prev_y.set, xscrollcommand=scroll_prev_x.set)
    tree_preview.pack(fill=tk.BOTH, expand=True)
    scroll_prev_y.config(command=tree_preview.yview)
    scroll_prev_x.config(command=tree_preview.xview)
    
    # Result Treeview
    scroll_res_y = tk.Scrollbar(tab_result)
    scroll_res_y.pack(side=tk.RIGHT, fill=tk.Y)
    scroll_res_x = tk.Scrollbar(tab_result, orient=tk.HORIZONTAL)
    scroll_res_x.pack(side=tk.BOTTOM, fill=tk.X)
    
    tree_result = ttk.Treeview(tab_result, yscrollcommand=scroll_res_y.set, xscrollcommand=scroll_res_x.set)
    tree_result.pack(fill=tk.BOTH, expand=True)
    scroll_res_y.config(command=tree_result.yview)
    scroll_res_x.config(command=tree_result.xview)
    
    root.mainloop()

if __name__ == "__main__":
    build_gui()
