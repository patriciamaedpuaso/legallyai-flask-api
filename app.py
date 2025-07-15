from flask import Flask, request, send_file
from flask_cors import CORS
from io import BytesIO
from html2docx import html2docx

app = Flask(__name__)
CORS(app, origins="*")

@app.route('/convert/html-to-docx', methods=['POST'])
def convert():
    data = request.get_json()
    html = data.get('html', '')

    docx_io = BytesIO()
    html2docx(html, docx_io)
    docx_io.seek(0)

    return send_file(
        docx_io,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='converted.docx'
    )

if __name__ == '__main__':
    app.run()
