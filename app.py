from flask import Flask, request, send_file, jsonify
from docx import Document
from flask_cors import CORS
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
from bs4 import BeautifulSoup

app = Flask(__name__)
CORS(app, origins="*")

# Helper functions
def parse_color(color_str):
    try:
        color_str = color_str.lstrip('#')
        return RGBColor(int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16))
    except:
        return None

def apply_styles(run, element):
    style = element.get('style', '')
    for rule in style.split(';'):
        if ':' in rule:
            key, value = [s.strip() for s in rule.split(':', 1)]
            if key == 'color':
                rgb = parse_color(value)
                if rgb:
                    run.font.color.rgb = rgb
            elif key == 'font-size':
                try:
                    size = int(value.replace('px', '').strip())
                    run.font.size = Pt(size)
                except:
                    pass
            elif key == 'font-family':
                run.font.name = value.split(',')[0].strip().strip('"\'')  # Take first font

    if element.name in ['strong', 'b']:
        run.bold = True
    if element.name in ['em', 'i']:
        run.italic = True
    if element.name == 'u':
        run.underline = True

def get_alignment(align_str):
    if align_str == 'center':
        return WD_ALIGN_PARAGRAPH.CENTER
    elif align_str == 'right':
        return WD_ALIGN_PARAGRAPH.RIGHT
    elif align_str == 'justify':
        return WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        return WD_ALIGN_PARAGRAPH.LEFT

def add_paragraph_with_formatting(document, element):
    if element.name and element.name.startswith('h'):
        level = int(element.name[1])
        p = document.add_paragraph(element.get_text(strip=True))
        p.style = f'Heading {min(level, 6)}'
        return

    if element.name == 'p':
        style = element.get('style', '')
        align = WD_ALIGN_PARAGRAPH.LEFT
        if 'text-align:' in style:
            try:
                align_val = style.lower().split('text-align:')[1].split(';')[0].strip()
                align = get_alignment(align_val)
            except:
                pass

        p = document.add_paragraph()
        p.alignment = align
        extract_text_with_formatting(p, element)

    elif element.name in ['ul', 'ol']:
        list_style = 'List Bullet' if element.name == 'ul' else 'List Number'
        for li in element.find_all('li', recursive=False):
            p = document.add_paragraph(style=list_style)
            extract_text_with_formatting(p, li)

@app.route('/convert/html-to-docx', methods=['POST'])
def convert_html_to_docx():
    data = request.get_json()
    html = data.get('html', '')

    document = Document()
    soup = BeautifulSoup(html, 'html.parser')

    body = soup.body if soup.body else soup

    for element in body.children:
        if getattr(element, 'name', None) in ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol']:
            add_paragraph_with_formatting(document, element)

    byte_io = BytesIO()
    document.save(byte_io)
    byte_io.seek(0)

    return send_file(byte_io, as_attachment=True, download_name='converted.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

if __name__ == '__main__':
    app.run(debug=True)
