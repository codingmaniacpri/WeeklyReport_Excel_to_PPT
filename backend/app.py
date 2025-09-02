from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import uuid
import traceback


# Import your updated processing logic
from excel_processing.read_excel import read_excel_sheets
from ppt_generation.slides import create_ppt_from_excel

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

UPLOAD_FOLDER = "./uploads"
GENERATED_FOLDER = "./generated"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)


@app.route("/api/upload-report", methods=["POST"])
def upload_report():
    # Check both files in request
    if 'excel' not in request.files or 'ppt' not in request.files:
        return jsonify({"message": "Both Excel and PPT files are required"}), 400

    excel_file = request.files['excel']
    ppt_file = request.files['ppt']

    if excel_file.filename == '' or ppt_file.filename == '':
        return jsonify({"message": "Invalid file(s)"}), 400

    # Generate unique ID for this request
    unique_id = str(uuid.uuid4())

    # Save Excel
    excel_upload_path = os.path.join(UPLOAD_FOLDER, f"{unique_id}.xlsx")
    excel_file.save(excel_upload_path)

    # Save PPT (this is the template)
    ppt_upload_path = os.path.join(UPLOAD_FOLDER, f"{unique_id}.pptx")
    ppt_file.save(ppt_upload_path)

    try:

        print("Excel path:", excel_upload_path)
        print("PPT path:", ppt_upload_path)

        # Step 1: Read all sheets from Excel
        sheets_data = read_excel_sheets(excel_upload_path)
        print("Sheets data keys:", sheets_data.keys())

        # Step 2: Generate PPT dynamically using the Excel data
        ppt_generated_path = os.path.join(GENERATED_FOLDER, f"{unique_id}_report.pptx")
        create_ppt_from_excel(ppt_upload_path, sheets_data, ppt_generated_path)
        print("PPT created:", ppt_generated_path)

        # Return download URL for only the PPT (Excel untouched)
        response = {
            "pptDownloadUrl": f"http://localhost:5000/api/download-report/{unique_id}/ppt"
        }
        return jsonify(response)

    except Exception as e:
        # Print the full traceback in the backend terminal
        traceback.print_exc()

        # Return full error message to the frontend (for now, to debug)
        return jsonify({
            "message": "Error processing report",
            "error": str(e),
            "traceback": traceback.format_exc()
        }), 500
        #return jsonify({"message": f"Error processing report: {str(e)}"}), 500


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

    return send_file(
        file_path,
        as_attachment=True,
        download_name=download_name,
        mimetype=mimetype
    )


if __name__ == "__main__":
    app.run(debug=True)
