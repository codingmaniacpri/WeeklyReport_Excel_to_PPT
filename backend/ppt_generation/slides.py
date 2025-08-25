from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import pandas as pd

def dataframe_to_ppt_with_template(df, template_file, output_file, rows_per_slide=12):
    prs = Presentation(template_file)

    layout_index = 1 if len(prs.slide_layouts) > 1 else 0
    content_layout = prs.slide_layouts[layout_index]

    columns = df.columns.tolist()
    data = df.values.tolist()

    insert_index = 1

    for i in range(0, len(data), rows_per_slide):
        slide = prs.slides.add_slide(content_layout)
        prs.slides._sldIdLst.insert(insert_index, prs.slides._sldIdLst[-1])
        insert_index += 1

        if slide.shapes.title:
            slide.shapes.title.text = "Dependencies"

        chunk = data[i:i+rows_per_slide]
        rows = len(chunk) + 1
        cols = len(columns)

        left, top, width, height = Inches(0.3), Inches(1.5), Inches(10), Inches(3)
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table

        # Header styling
        for j, col_name in enumerate(columns):
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
            cell.fill.fore_color.rgb = RGBColor(189, 215, 238)

        # Data rows
        for r, row in enumerate(chunk, start=1):
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
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    prs.save(output_file)
    return output_file
