from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pypandoc
import tempfile
import os

app = Flask(__name__)
CORS(app, origins="*")  # Allow all origins for CORS

# Path to your custom reference DOCX (must be in same folder or provide full path)
REFERENCE_DOCX_PATH = os.path.join(os.path.dirname(__file__), 'reference.docx')

@app.route('/convert/html-to-docx', methods=['POST'])
def convert_html_to_docx():
    try:
        data = request.get_json()
        html = data.get('html', '')

        if not html.strip():
            return jsonify({"error": "Empty HTML content"}), 400

        # Wrap HTML in <html><body> to ensure valid structure
        full_html = f"<html><body>{html}</body></html>"

        # Create temporary file for output DOCX
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            output_path = tmp_docx.name

        # Convert HTML to DOCX using Pandoc and reference file
        pypandoc.convert_text(
            full_html,
            to='docx',
            format='html',
            outputfile=output_path,
            extra_args=[
                f'--reference-doc={REFERENCE_DOCX_PATH}',
                '--standalone'
            ]
        )

        # Read and return file as attachment
        with open(output_path, 'rb') as f:
            docx_data = f.read()

        os.remove(output_path)  # Clean up temp file

        return send_file(
            BytesIO(docx_data),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='converted.docx'
        )

    except Exception as e:
        return jsonify({"error": f"Conversion failed: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True)
