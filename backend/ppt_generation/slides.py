from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import pandas as pd

def style_table(table, df, start_row, rows_per_slide):
    """
    Apply consistent styling for header and data rows in the PPT table.
    """
    # ✅ Header styling
    for j, col_name in enumerate(df.columns):
        cell = table.cell(0, j)
        cell.text_frame.clear()
        p = cell.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = str(col_name)

        font = run.font
        font.name = "Calibri"
        font.bold = True
        font.size = Pt(14)
        font.color.rgb = RGBColor(0, 0, 0)

        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(189, 215, 238)  # Light blue

    # ✅ Data rows
    for r, row in enumerate(df.iloc[start_row:start_row + rows_per_slide].values, start=1):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text_frame.clear()
            p = cell.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = str(val) if pd.notna(val) else ""

            font = run.font
            font.name = "Calibri"
            font.size = Pt(12)
            font.color.rgb = RGBColor(0, 0, 0)

            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White


def create_ppt_from_excel(template_path, excel_data, output_path, rows_per_slide=12):
    """
    Creates a PPT from Excel sheet data (dict of {sheet_name: dataframe}).
    Keeps the first slide of template as-is, starts inserting data from slide 2.
    """
    # Load template
    prs = Presentation(template_path)

    for sheet_name, df in excel_data.items():

        # ✅ Skip sheets with no data or no columns
        if df.empty or len(df.columns) == 0:
            continue

        total_rows = len(df)
        start_row = 0
        slide_count = 1

        while start_row < total_rows:
            # Use Title and Content layout (usually index 1)
            slide_layout = prs.slide_layouts[1]    # Title + Content
            slide = prs.slides.add_slide(slide_layout)

            # Title with (Contd..) if multiple slides for the same sheet
            title_text = f"{sheet_name} (Contd..)" if slide_count > 1 else sheet_name
            
            if slide.shapes.title:
                slide.shapes.title.text = title_text
            else:
                left, top, width, height = Inches(0.5), Inches(0.3), Inches(9), Inches(1)
                textbox = slide.shapes.add_textbox(left, top, width, height)
                textbox.text_frame.text = title_text

            # Add table
            rows = min(rows_per_slide, total_rows - start_row) + 1  # +1 for header
            cols = len(df.columns)
            left = Inches(0.5)
            top = Inches(1.5)
            width = Inches(9)
            height = Inches(0.8 * rows)  # dynamic height

            table = slide.shapes.add_table(rows, cols, left, top, width, height).table

            # Apply styling
            style_table(table, df, start_row, rows_per_slide)

            # Move to next chunk of rows
            start_row += rows_per_slide
            slide_count += 1

    # Save output PPT
    prs.save(output_path)
    return output_path


# from pptx import Presentation
# from pptx.util import Pt, Inches
# from pptx.dml.color import RGBColor
# import pandas as pd

# def hex_to_rgb(color_obj):
#     """
#     Convert openpyxl color (string, RGB object, or None) to RGB tuple.
#     """
#     if color_obj is None:
#         return (0, 0, 0)

#     # Case 1: Already a string like "FF0000" or "0xFF0000"
#     if isinstance(color_obj, str):
#         hex_str = color_obj.replace("0x", "").replace("#", "")
#         if len(hex_str) == 8:  # ARGB
#             hex_str = hex_str[2:]
#         return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))

#     # Case 2: openpyxl Color/RGB object
#     if hasattr(color_obj, "rgb") and color_obj.rgb:
#         hex_str = color_obj.rgb.replace("0x", "").replace("#", "")
#         if len(hex_str) == 8:
#             hex_str = hex_str[2:]
#         return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))

#     # Case 3: Theme or other → fallback
#     return (0, 0, 0)


# def style_table(table, df, styles, start_row, rows_per_slide):
#     """
#     Apply Excel-like styling to the PPT table cells.
#     """
#     # --- Header styling ---
#     for j, col_name in enumerate(df.columns):
#         cell = table.cell(0, j)
#         cell.text_frame.clear()
#         run = cell.text_frame.paragraphs[0].add_run()
#         run.text = str(col_name)

#         font = run.font
#         font.name = "Calibri"
#         font.bold = True
#         font.size = Pt(14)
#         font.color.rgb = RGBColor(0, 0, 0)
#         cell.fill.solid()
#         cell.fill.fore_color.rgb = RGBColor(189, 215, 238)

#     # --- Data rows with Excel styling ---
#     for r, row in enumerate(df.iloc[start_row:start_row + rows_per_slide].values, start=1):
#         for c, val in enumerate(row):
#             cell = table.cell(r, c)
#             cell.text_frame.clear()
#             run = cell.text_frame.paragraphs[0].add_run()
#             run.text = str(val) if pd.notna(val) else ""

#             # Default font
#             font = run.font
#             font.name = "Calibri"
#             font.size = Pt(12)
#             font.color.rgb = RGBColor(0, 0, 0)
#             cell.fill.solid()
#             cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

#             # --- Apply extracted Excel styles ---
#             style = styles.get((start_row + r, c))  # row index adjusted
#             if style:
#                 if style["font_name"]:
#                     font.name = style["font_name"]
#                 if style["font_size"]:
#                     font.size = Pt(style["font_size"])
#                 if style["bold"]:
#                     font.bold = True
#                 if style["italic"]:
#                     font.italic = True
#                 if style["color"]:
#                     rgb = hex_to_rgb(style["color"])
#                     font.color.rgb = RGBColor(*rgb)
#                 if style["fill_color"]:
#                     rgb = hex_to_rgb(style["fill_color"])
#                     cell.fill.solid()
#                     cell.fill.fore_color.rgb = RGBColor(*rgb)


# def create_ppt_from_excel(template_path, excel_data, output_path, rows_per_slide=12):
#     """
#     Create PPT with Excel values and formatting.
#     """
#     prs = Presentation(template_path)

#     for sheet_name, content in excel_data.items():
#         df = content["data"]
#         styles = content["styles"]

#         if df.empty or len(df.columns) == 0:
#             continue

#         total_rows = len(df)
#         start_row = 0
#         slide_count = 1

#         while start_row < total_rows:
#             slide_layout = prs.slide_layouts[1]
#             slide = prs.slides.add_slide(slide_layout)

#             title_text = f"{sheet_name} (Contd..)" if slide_count > 1 else sheet_name
#             if slide.shapes.title:
#                 slide.shapes.title.text = title_text

#             rows = min(rows_per_slide, total_rows - start_row) + 1
#             cols = len(df.columns)
#             left, top, width, height = Inches(0.5), Inches(1.5), Inches(9), Inches(0.8 * rows)
#             table = slide.shapes.add_table(rows, cols, left, top, width, height).table

#             style_table(table, df, styles, start_row, rows_per_slide)

#             start_row += rows_per_slide
#             slide_count += 1

#     prs.save(output_path)
#     return output_path
