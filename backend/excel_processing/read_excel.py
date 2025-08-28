import pandas as pd
import re

def read_excel_sheets(file_path):
    """
    Reads all sheets in the Excel file and returns a dictionary 
    {sheet_name: dataframe}.
    """
    # Load all sheets into a dictionary
    sheets = pd.read_excel(file_path, sheet_name=None, header=0)

    # Optional cleanup: remove unnamed index columns if they exist
    cleaned_sheets = {}
    for name, df in sheets.items():
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        cleaned_sheets[name] = df

    return cleaned_sheets