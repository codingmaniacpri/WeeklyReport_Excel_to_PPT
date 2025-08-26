# from flask import Flask, request, jsonify, send_file
# from flask_cors import CORS
# import os
# import uuid

# from utils.file_utils import save_uploaded_file, get_file_path
# from excel_processing.read_excel import get_latest_date_comments
# from ppt_generation.slides import dataframe_to_ppt_with_template

# UPLOAD_FOLDER = "static/uploads"
# OUTPUT_FOLDER = "static/outputs"

# app = Flask(__name__)
# CORS(app)

# app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
# app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER

# # ensure dirs exist
# os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# @app.route("/upload_excel", methods=["POST"])
# def upload_files():
#     """
#     Endpoint: /upload
#     Accepts Excel + PPT template file uploads
#     Generates new PPT and returns download link
#     """
#     if "excel_file" not in request.files or "ppt_template" not in request.files:
#         return jsonify({"error": "Both excel_file and ppt_template files are required"}), 400

#     excel_file = request.files["excel_file"]
#     template_file = request.files["ppt_template"]

#     # Save files
#     excel_path = save_uploaded_file(excel_file, app.config["UPLOAD_FOLDER"])
#     template_path = save_uploaded_file(template_file, app.config["UPLOAD_FOLDER"])

#     # Process Excel
#     df, _ = get_latest_date_comments(excel_path)

#     # Generate PPT
#     output_filename = f"generated_{uuid.uuid4().hex}.pptx"
#     output_path = os.path.join(app.config["OUTPUT_FOLDER"], output_filename)

#     dataframe_to_ppt_with_template(df, template_path, output_path)

#     # Return plain download URL string compatible with frontend's XHR usage
#     return f"http://localhost:5000/download/{output_filename}", 200

#     # return jsonify({
#     #     "message": "PPT generated successfully",
#     #     "download_url": f"/download/{output_filename}"
#     # })


# @app.route("/download/<filename>", methods=["GET"])
# def download_file(filename):
#     """
#     Download generated PPT
#     """
#     file_path = os.path.join(app.config["OUTPUT_FOLDER"], filename)
#     if not os.path.exists(file_path):
#         return jsonify({"error": "File not found"}), 404
#     return send_file(file_path, as_attachment=True)


# if __name__ == "__main__":
#     app.run(debug=True, port=5000)



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
 