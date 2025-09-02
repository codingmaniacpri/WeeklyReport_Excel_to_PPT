import pandas as pd
import openpyxl

def detect_header(df, max_rows=10, max_cols=10):
    """
    Detect if the sheet has a proper header or not.
    Looks at the first `max_rows` rows and `max_cols` columns to find data.
    
    Returns:
        cleaned_df (pd.DataFrame): DataFrame with correct header handling
    """
    # Drop completely empty rows and columns first
    df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)

    if df.empty:
        return pd.DataFrame()  # return empty if no data at all

    # Look at a subset (first few rows & columns) for detecting headers
    subset = df.iloc[:max_rows, :max_cols]

    # Case 1: If first row has mostly text and not numeric → assume it’s header
    first_row = subset.iloc[0].astype(str)
    text_ratio = sum(first_row.str.strip() != "") / len(first_row)

    if text_ratio > 0.5:  
        # Likely that the first row is a header
        df.columns = df.iloc[0].astype(str)
        df = df.drop(df.index[0])
    else:
        # Otherwise, just assign generic column names
        df.columns = [f"Column_{i}" for i in range(1, len(df.columns) + 1)]

    # Final cleanup again
    df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)

    return df.reset_index(drop=True)


def read_excel_sheets(file_path):
    """
    Reads only visible and non-empty sheets in the Excel file 
    and returns a dictionary {sheet_name: dataframe}.
    """
    # Load workbook with openpyxl to check sheet visibility
    wb = openpyxl.load_workbook(file_path, read_only=False)
    visible_sheets = [sheet.title for sheet in wb.worksheets if sheet.sheet_state == "visible"]

    print(f"Visible sheets found: {visible_sheets}")  # Debugging

    # Read all visible sheets without headers first
    sheets = pd.read_excel(file_path, sheet_name=visible_sheets, header=None)

    cleaned_sheets = {}
    for name, df in sheets.items():
        cleaned_df = detect_header(df)

        if not cleaned_df.empty:
            print(f"Processing sheet: {name} (rows: {len(cleaned_df)})")     # Debugging line
            cleaned_sheets[name] = cleaned_df
        else:
            print(f"Skipping empty sheet: {name}")

    return cleaned_sheets


# import pandas as pd
# import openpyxl
# from openpyxl.styles import Font, PatternFill

# def detect_header(df, max_rows=10, max_cols=10):
#     """
#     Detect if the sheet has a proper header or not.
#     """
#     df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)

#     if df.empty:
#         return pd.DataFrame()

#     subset = df.iloc[:max_rows, :max_cols]
#     first_row = subset.iloc[0].astype(str)
#     text_ratio = sum(first_row.str.strip() != "") / len(first_row)

#     if text_ratio > 0.5:
#         df.columns = df.iloc[0].astype(str)
#         df = df.drop(df.index[0])
#     else:
#         df.columns = [f"Column_{i}" for i in range(1, len(df.columns) + 1)]

#     df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
#     return df.reset_index(drop=True)


# def read_excel_sheets(file_path):
#     """
#     Reads visible sheets in Excel, extracting both data (pandas df) and styles (dict).
#     Returns: {sheet_name: {"data": df, "styles": styles}}
#     """
#     wb = openpyxl.load_workbook(file_path, read_only=False, data_only=True)
#     visible_sheets = [sheet for sheet in wb.worksheets if sheet.sheet_state == "visible"]

#     result = {}

#     for sheet in visible_sheets:
#         # --- Extract raw values into dataframe ---
#         data = sheet.values
#         df = pd.DataFrame(data)
#         cleaned_df = detect_header(df)

#         if cleaned_df.empty:
#             continue

#         # --- Extract styles ---
#         styles = {}
#         for row in sheet.iter_rows():
#             for cell in row:
#                 if cell.row == 1:  # skip column headers row before cleaning
#                     continue
#                 if cell.value is None:
#                     continue
#                 styles[(cell.row - 1, cell.column - 1)] = {
#                     "font_name": cell.font.name,
#                     "font_size": cell.font.sz,
#                     "bold": cell.font.bold,
#                     "italic": cell.font.italic,
#                     "color": (cell.font.color.rgb if cell.font.color else "000000"),
#                     "fill_color": (cell.fill.fgColor.rgb if isinstance(cell.fill, PatternFill) else None),
#                 }

#         result[sheet.title] = {
#             "data": cleaned_df,
#             "styles": styles
#         }

#     return result
