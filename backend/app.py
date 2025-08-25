from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import uuid

from utils.file_utils import save_uploaded_file, get_file_path
from excel_processing.read_excel import get_latest_date_comments
from ppt_generation.slides import dataframe_to_ppt_with_template

UPLOAD_FOLDER = "static/uploads"
OUTPUT_FOLDER = "static/outputs"

app = Flask(__name__)
CORS(app)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER

# ensure dirs exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/upload", methods=["POST"])
def upload_files():
    """
    Endpoint: /upload
    Accepts Excel + PPT template file uploads
    Generates new PPT and returns download link
    """
    if "excel" not in request.files or "template" not in request.files:
        return jsonify({"error": "Both excel and template files are required"}), 400

    excel_file = request.files["excel"]
    template_file = request.files["template"]

    # Save files
    excel_path = save_uploaded_file(excel_file, app.config["UPLOAD_FOLDER"])
    template_path = save_uploaded_file(template_file, app.config["UPLOAD_FOLDER"])

    # Process Excel
    df, _ = get_latest_date_comments(excel_path)

    # Generate PPT
    output_filename = f"generated_{uuid.uuid4().hex}.pptx"
    output_path = os.path.join(app.config["OUTPUT_FOLDER"], output_filename)

    dataframe_to_ppt_with_template(df, template_path, output_path)

    return jsonify({
        "message": "PPT generated successfully",
        "download_url": f"/download/{output_filename}"
    })


@app.route("/download/<filename>", methods=["GET"])
def download_file(filename):
    """
    Download generated PPT
    """
    file_path = os.path.join(app.config["OUTPUT_FOLDER"], filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404
    return send_file(file_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
