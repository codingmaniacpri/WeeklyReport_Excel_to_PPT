from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import pandas as pd

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
            
            # If no title placeholder, add a textbox manually
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

            # Header row
            for col_idx, col_name in enumerate(df.columns):
                cell = table.cell(0, col_idx)
                cell.text = str(col_name)
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.bold = True
                        run.font.size = Pt(12)
                        run.font.color.rgb = RGBColor(255, 255, 255)
                #cell.text_frame.paragraphs[0].font.bold = True
                cell.fill.solid()
                #cell.fill.fore_color.rgb = RGBColor(79, 129, 189)  # Blue header
                cell.fill.fore_color.rgb = RGBColor(0, 51, 102)
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

            # # Fill table rows
            # for i in range(1, rows):
            #     for j in range(cols):
            #         val = df.iloc[start_row + i - 1, j]
            #         table.cell(i, j).text = str(val) if not pd.isna(val) else ""

            # start_row += (rows - 1)
            # slide_count += 1

            # Fill rows
            for row_idx, row_data in enumerate(df.iloc[start_row:start_row + rows_per_slide].values, start=1):
                for col_idx, value in enumerate(row_data):
                    table.cell(row_idx, col_idx).text = str(value) if not pd.isna(value) else ""

            start_row += rows_per_slide
            slide_count += 1

    # Save output PPT
    prs.save(output_path)
    return output_path






