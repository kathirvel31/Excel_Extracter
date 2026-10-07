def export_to_excel(df, save_path):
    """
    Exports a pandas DataFrame to an Excel file.
    """
    try:
        df.to_excel(save_path, index=False)
    except Exception as e:
        raise Exception(f"Failed to export data to {save_path}: {str(e)}")
