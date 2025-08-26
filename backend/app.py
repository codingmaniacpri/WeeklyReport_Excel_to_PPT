from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import uuid
import shutil

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

UPLOAD_FOLDER = "./uploads"
GENERATED_FOLDER = "./generated"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)


@app.route("/api/upload-report", methods=["POST"])
def upload_report():
    if 'file' not in request.files:
        return jsonify({"message": "No file part"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"message": "No selected file"}), 400

    # Save the uploaded file with a unique name
    unique_id = str(uuid.uuid4())
    upload_path = os.path.join(UPLOAD_FOLDER, f"{unique_id}.xlsx")
    file.save(upload_path)

    # PROCESS THE FILE HERE
    # For demo, just copy upload as 'generated' report
    generated_path = os.path.join(GENERATED_FOLDER, f"{unique_id}_report.xlsx")
    shutil.copyfile(upload_path, generated_path)

    download_url = f"http://localhost:5000/api/download-report/{unique_id}"

    return jsonify({"downloadUrl": download_url})


@app.route("/api/download-report/<file_id>", methods=["GET"])
def download_report(file_id):
    generated_path = os.path.join(GENERATED_FOLDER, f"{file_id}_report.xlsx")
    if not os.path.exists(generated_path):
        return jsonify({"message": "Report not found"}), 404

    return send_file(generated_path,
                     as_attachment=True,
                     download_name="report.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(debug=True)
