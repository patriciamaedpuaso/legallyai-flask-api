from flask import Flask, request, send_file
from flask_cors import CORS
from io import BytesIO
import pypandoc
import tempfile
import os
from bs4 import BeautifulSoup

app = Flask(__name__)
CORS(app, origins="*")

def clean_html(html):
    soup = BeautifulSoup(html, 'html.parser')

    # Convert <p style="text-align:..."> to <p align="...">
    for p in soup.find_all('p'):
        if 'style' in p.attrs:
            style = p['style']
            if 'text-align' in style:
                align = style.split('text-align:')[1].split(';')[0].strip()
                p['align'] = align
                del p['style']

    # Convert <span style="color:..."> to <font color="...">
    for span in soup.find_all('span'):
        if 'style' in span.attrs and 'color' in span['style']:
            color = span['style'].split('color:')[1].split(';')[0].strip()
            font_tag = soup.new_tag('font', color=color)
            font_tag.string = span.get_text()
            span.replace_with(font_tag)

    # Convert <em style="color:..."> to <font color="..."><em>...</em></font>
    for em in soup.find_all('em'):
        if 'style' in em.attrs and 'color' in em['style']:
            color = em['style'].split('color:')[1].split(';')[0].strip()
            font_tag = soup.new_tag('font', color=color)
            em_copy = soup.new_tag('em')
            em_copy.string = em.get_text()
            font_tag.append(em_copy)
            em.replace_with(font_tag)

    return str(soup)

@app.route('/convert/html-to-docx', methods=['POST'])
def convert():
    data = request.get_json()
    html = data.get('html', '')

    try:
        cleaned_html = clean_html(html)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            output_path = tmp.name
            pypandoc.convert_text(
                cleaned_html,
                'docx',
                format='html',
                outputfile=output_path
            )

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
