from pptx import Presentation

def load_template(path: str):
    return Presentation(path)

def save_presentation(prs, out_path: str):
    prs.save(out_path)
