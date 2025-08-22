from pptx.util import Inches

def add_title_slide(prs, company_name, logo_path=None):
    """
    Add first slide with company name and optional logo.
    """
    slide_layout = prs.slide_layouts[0]  # Title Slide layout
    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title
    title.text = company_name

    if logo_path:
        slide.shapes.add_picture(logo_path, Inches(6), Inches(0.5), height=Inches(1))
    return slide

def add_table_slide(prs, title_text, table_data):
    """
    Add a slide with a title and a table (table_data is a list of lists).
    """
    slide_layout = prs.slide_layouts[5]  # Title and Content layout
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title_text

    rows = len(table_data)
    cols = len(table_data) if rows > 0 else 0

    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(9)
    height = Inches(0.8)

    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    for i, row in enumerate(table_data):
        for j, val in enumerate(row):
            table.cell(i, j).text = str(val)

    return slide
