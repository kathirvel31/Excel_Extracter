import pandas as pd

def get_sheet_names(file_path):
    """
    Reads an Excel file and returns a list of its sheet names.
    """
    try:
        xl = pd.ExcelFile(file_path)
        return xl.sheet_names
    except Exception as e:
        raise Exception(f"Failed to read Excel file sheets: {str(e)}")

def load_sheet_data(file_path, sheet_name):
    """
    Loads a specific sheet from an Excel file into a pandas DataFrame.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df
    except Exception as e:
        raise Exception(f"Failed to load sheet '{sheet_name}': {str(e)}")
