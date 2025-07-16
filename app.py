from flask import Flask, request, send_file
from flask_cors import CORS
from io import BytesIO
import pypandoc
import tempfile
import os

app = Flask(__name__)
CORS(app, origins="*")

@app.route('/convert/html-to-docx', methods=['POST'])
def convert():
    data = request.get_json()
    html = data.get('html', '')

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            output_path = tmp.name
            pypandoc.convert_text(html, 'docx', format='html', outputfile=output_path)
        
        with open(output_path, 'rb') as f:
            docx_data = f.read()

        os.remove(output_path)

        return send_file(
            BytesIO(docx_data),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='converted.docx'
        )
    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    app.run()
