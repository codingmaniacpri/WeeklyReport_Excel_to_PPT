import os
import uuid

UPLOAD_FOLDER = 'static/uploads/'

def save_file(file_storage):
    """
    Save uploaded file with a unique name to UPLOAD_FOLDER.
    Returns the saved file path.
    """
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)

    ext = os.path.splitext(file_storage.filename)[1]
    unique_name = str(uuid.uuid4()) + ext
    file_path = os.path.join(UPLOAD_FOLDER, unique_name)
    file_storage.save(file_path)
    return file_path
