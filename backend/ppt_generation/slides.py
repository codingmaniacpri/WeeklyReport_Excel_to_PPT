from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
import pandas as pd
import datetime


def set_cell_border(cell, border_color="000000", border_width="12700"):
    """
    Add border to a pptx table cell.
    Default: Black, ~0.5pt width.
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for line in ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]:
        ln = tcPr.find(qn(line))
        if ln is None:
            ln = parse_xml(
                f'<{line} w="{border_width}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                f'<a:solidFill><a:srgbClr val="{border_color}"/></a:solidFill></{line}>'
            )
            tcPr.append(ln)


def style_header_cell(cell, font_size=12):
    """
    Apply styling to a header cell.
    """
    cell.text_frame.word_wrap = True  # ✅ enable wrapping
    for p in cell.text_frame.paragraphs:
        for run in p.runs:
            run.font.bold = True
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor(255, 255, 255)  # White text
        p.alignment = PP_ALIGN.CENTER

    # Background + border
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(0, 51, 102)  # Dark blue
    set_cell_border(cell)


def format_value(value):
    if pd.isna(value):
        return ""
    # Handle Python datetime.datetime or pandas Timestamp objects
    if isinstance(value, (datetime.datetime, datetime.date, pd.Timestamp)):
        # For datetime.datetime convert to date to remove time part
        if isinstance(value, datetime.datetime):
            value = value.date()
        return value.strftime("%d-%m-%Y")  # Format as DD-MM-YYYY
    # If value is string that contains date-time, parse and format
    elif isinstance(value, str):
        try:
            # Try to parse string as date-time with time part, then format
            parsed = datetime.datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
            return parsed.strftime("%d-%m-%Y")
        except Exception:
            # If parsing fails, return original string
            return value
    return str(value)


def style_data_cell(cell, value, font_size=10):
    """
    Apply styling to a data cell.
    """
    cell.text_frame.text = format_value(value)
    cell.text_frame.word_wrap = True  # ✅ enable wrapping

    for p in cell.text_frame.paragraphs:
        for run in p.runs:
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor(0, 0, 0)  # Black text
        p.alignment = PP_ALIGN.LEFT

    # Background + border
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White
    set_cell_border(cell)
    
def adjust_row_height(table, row_idx, df_row, col_widths, base_height=0.4):
    """
    Adjust row height dynamically based on text length and column width.
    """
    max_lines = 1
    for col_idx, value in enumerate(df_row):
        if col_idx >= len(col_widths):
            continue  # skip if more columns than widths
        
        text = format_value(value)
        approx_chars_per_line = int(col_widths[col_idx].inches * 12)  # ~12 chars per inch
        line_count = (len(text) // approx_chars_per_line) + 1
        if line_count > max_lines:
            max_lines = line_count

    table.rows[row_idx].height = Inches(base_height * max_lines)




def create_ppt_from_excel(template_path, excel_data, output_path, rows_per_slide=5):
    """
    Creates a PPT from Excel sheet data (dict of {sheet_name: dataframe}).
    Keeps the first slide of template as-is, starts inserting data from slide 2.
    """
    prs = Presentation(template_path)

    for sheet_name, df in excel_data.items():

        if df.empty or len(df.columns) == 0:
            continue

        total_rows = len(df)
        start_row = 0
        slide_count = 1

        while start_row < total_rows:
            slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(slide_layout)

            # Title
            title_text = f"{sheet_name} (Contd..)" if slide_count > 1 else sheet_name
            if slide.shapes.title:
                slide.shapes.title.text = title_text
            else:
                left, top, width, height = Inches(0.5), Inches(0.3), Inches(9), Inches(1)
                textbox = slide.shapes.add_textbox(left, top, width, height)
                textbox.text_frame.text = title_text

            # Table
            rows = min(rows_per_slide, total_rows - start_row) + 1
            cols = len(df.columns)
            left, top, width, height = Inches(0.5), Inches(1.5), Inches(9), Inches(0.8 * rows)

            table = slide.shapes.add_table(rows, cols, left, top, width, height).table

            # Column widths (customize as needed)
            col_widths = [
                Inches(0.6), Inches(1.5), Inches(2.5), Inches(1.2), Inches(1.2),
                Inches(1.2), Inches(1.5), Inches(0.6), Inches(3.5)
            ]
            for i, w in enumerate(col_widths[:cols]):
                table.columns[i].width = w

            # Header row
            for col_idx, col_name in enumerate(df.columns):
                cell = table.cell(0, col_idx)
                cell.text = str(col_name)
                style_header_cell(cell)

            # Data rows
            for row_idx, row_data in enumerate(
                df.iloc[start_row:start_row + rows_per_slide].values, start=1
            ):
                for col_idx, value in enumerate(row_data):
                    style_data_cell(table.cell(row_idx, col_idx), value)
                    
                # ✅ adjust height after filling row
                adjust_row_height(table, row_idx, row_data, col_widths)

            start_row += rows_per_slide
            slide_count += 1

    prs.save(output_path)
    return output_path
