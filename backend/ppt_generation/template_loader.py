from pptx import Presentation

def load_ppt_template(file_stream):
    """
    Load PPTX presentation from file-like object.
    """
    prs = Presentation(file_stream)
    return prs
