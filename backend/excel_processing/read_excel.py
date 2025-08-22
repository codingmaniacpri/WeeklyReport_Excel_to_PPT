import pandas as pd

def read_excel_file(file_stream):
    """
    Read Excel file from file-like object.
    Returns dict of sheet_name -> DataFrame
    """
    excel_sheets = pd.read_excel(file_stream, sheet_name=None)
    return excel_sheets
