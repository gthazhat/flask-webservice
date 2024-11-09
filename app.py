# app.py

from flask import Flask, jsonify
import subprocess

app = Flask(__name__)


# Endpoint to run the scripts
@app.route('/run_automation', methods=['GET'])
def run_automation():
    try:
        # Run Cp360_CompleteAutomation.py, which should call the other scripts
        result = subprocess.run(['python', 'Cp360_CompleteAutomation.py'], capture_output=True, text=True)

        # Capture output and errors
        output = result.stdout
        error = result.stderr

        if result.returncode == 0:
            return jsonify({"status": "success", "output": output}), 200
        else:
            return jsonify({"status": "error", "error": error}), 500

    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
