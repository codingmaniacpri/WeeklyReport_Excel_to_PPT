from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import uuid
import traceback
import json
import pandas as pd

from excel_processing.read_excel import read_visible_excel_sheets_to_json
from ppt_generation.slides import create_ppt_from_excel  # expects dict with sheet data

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "./uploads"
GENERATED_FOLDER = "./generated"
EXTRACTED_FOLDER = "./extracted_json"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)
os.makedirs(EXTRACTED_FOLDER, exist_ok=True)

@app.route("/api/upload-report", methods=["POST"])
def upload_report():
    if 'excel' not in request.files or 'ppt' not in request.files:
        return jsonify({"message": "Both Excel and PPT files are required"}), 400
    excel_file = request.files['excel']
    ppt_file = request.files['ppt']
    if excel_file.filename == '' or ppt_file.filename == '':
        return jsonify({"message": "Invalid file(s)"}), 400

    unique_id = str(uuid.uuid4())
    excel_upload_path = os.path.join(UPLOAD_FOLDER, f"{unique_id}.xlsx")
    ppt_upload_path = os.path.join(UPLOAD_FOLDER, f"{unique_id}.pptx")
    excel_file.save(excel_upload_path)
    ppt_file.save(ppt_upload_path)

    try:
        print("Excel path:", excel_upload_path)
        print("PPT path:", ppt_upload_path)

        # Step 1: Extract sheets as styled JSON files
        extracted_json_files = read_visible_excel_sheets_to_json(excel_upload_path, EXTRACTED_FOLDER)
        print("Extracted JSON files:", extracted_json_files)

        # Step 2: Load JSON files into dictionary to pass to PPT creator
        sheets_data = {}
        for json_path in extracted_json_files:
            with open(json_path, "r", encoding="utf-8") as f:
                sheet_json = json.load(f)  # list of dicts with nested 'value' and 'style'
            
            # Extract only the "value" from each cell for DataFrame construction
            clean_rows = []
            for row in sheet_json:
                clean_row = {col: cell_info["value"] if isinstance(cell_info, dict) else cell_info
                             for col, cell_info in row.items()}
                clean_rows.append(clean_row)

            # Convert to DataFrame
            df = pd.DataFrame(clean_rows)
            sheets_data[os.path.splitext(os.path.basename(json_path))[0]] = df

        print("Sheets data loaded:", sheets_data.keys())

        # Step 3: Generate PPT using template & sheets_data dict
        ppt_generated_path = os.path.join(GENERATED_FOLDER, f"{unique_id}_report.pptx")
        create_ppt_from_excel(ppt_upload_path, sheets_data, ppt_generated_path)

        if not os.path.exists(ppt_generated_path):
            return jsonify({"message": "Failed to create PPT file"}), 500

        response = {
            "pptDownloadUrl": f"http://localhost:5000/api/download-report/{unique_id}/ppt",
            "extractedJson": [os.path.basename(j) for j in extracted_json_files]
        }
        return jsonify(response)

    except Exception as e:
        traceback.print_exc()
        return jsonify({
            "message": "Error processing report",
            "error": str(e),
            "traceback": traceback.format_exc()
        }), 500

@app.route("/api/download-report/<file_id>/<file_type>", methods=["GET"])
def download_report(file_id, file_type):
    if file_type == "ppt":
        file_path = os.path.join(GENERATED_FOLDER, f"{file_id}_report.pptx")
        download_name = "PptReport.pptx"
        mimetype = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    else:
        return jsonify({"message": "Invalid file type"}), 400

    if not os.path.exists(file_path):
        return jsonify({"message": "Report not found"}), 404

    return send_file(file_path, as_attachment=True, download_name=download_name, mimetype=mimetype)

if __name__ == "__main__":
    app.run(debug=True)
