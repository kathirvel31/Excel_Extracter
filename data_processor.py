import pandas as pd

def process_data(df, selected_columns, sort_order="A-Z"):
    """
    Extracts selected columns from the dataframe, removes rows with empty/null 
    values in those columns, and sorts the data alphabetically.
    """
    if not selected_columns:
        return pd.DataFrame()
    
    try:
        # Extract only the selected columns
        df_filtered = df[selected_columns].copy()
        
        # Remove rows where ALL selected columns are empty/null.
        df_filtered.dropna(subset=selected_columns, how='all', inplace=True)
        
        # Safely convert columns to object type and fill NaNs with empty string
        for col in df_filtered.columns:
            # fillna first, or astype first?
            # astype('object') forces the column to accept mixed types (str and float)
            df_filtered[col] = df_filtered[col].astype('object')
            df_filtered[col] = df_filtered[col].fillna("")
            
        # Sort alphabetically
        ascending = True if sort_order == "A-Z" else False
        
        # Sort by all selected columns sequentially, making it case-insensitive
        df_filtered.sort_values(
            by=selected_columns, 
            ascending=ascending, 
            inplace=True, 
            key=lambda col: col.astype(str).str.lower()
        )
        
        return df_filtered
    except Exception as e:
        raise Exception(f"Error processing data: {str(e)}")
