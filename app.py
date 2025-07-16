from flask import Flask, request, send_file
from flask_cors import CORS
from io import BytesIO
import pypandoc
import tempfile
import os
from bs4 import BeautifulSoup

app = Flask(__name__)
CORS(app, origins="*")

def preprocess_html(html: str) -> str:
    soup = BeautifulSoup(html, 'html.parser')

    for tag in soup.find_all(True):
        style = tag.get('style', '')

        # Handle text alignment
        if 'text-align:center' in style:
            tag['align'] = 'center'
        elif 'text-align:right' in style:
            tag['align'] = 'right'
        elif 'text-align:left' in style:
            tag['align'] = 'left'

        # Handle text color
        if 'color:' in style:
            color_value = style.split('color:')[1].split(';')[0].strip()
            tag['style'] = f'color:{color_value}'  # Keep only color
        else:
            tag.attrs.pop('style', None)  # Remove unused style

    return str(soup)

@app.route('/convert/html-to-docx', methods=['POST'])
def convert():
    data = request.get_json()
    html = data.get('html', '')

    try:
        processed_html = preprocess_html(html)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            output_path = tmp.name
            pypandoc.convert_text(processed_html, 'docx', format='html', outputfile=output_path)

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
