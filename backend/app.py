from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io

app = Flask(__name__)
CORS(app)

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    # Placeholder - actual processing to be implemented
    return jsonify({"message": "API is running. Implement PPT generation logic."})

if __name__ == '__main__':
    app.run(debug=True)
