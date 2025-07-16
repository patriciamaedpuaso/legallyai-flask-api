from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from io import BytesIO
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

app = Flask(__name__)
CORS(app, origins="*")


def hex_to_rgb(hex_color):
    """Convert hex color string to RGBColor"""
    hex_color = hex_color.lstrip('#')
    r, g, b = [int(hex_color[i:i+2], 16) for i in (0, 2, 4)]
    return RGBColor(r, g, b)


@app.route('/convert/html-to-docx', methods=['POST'])
def convert_html_to_docx():
    data = request.get_json()
    html = data.get("html", "")

    # Wrap in <html><body> if not present
    if "<html" not in html:
        html = f"<html><body>{html}</body></html>"

    soup = BeautifulSoup(html, "html.parser")
    document = Document()

    for p in soup.find_all("p"):
        style = p.get("style", "")
        align_match = re.search(r"text-align\s*:\s*(\w+)", style)
        align = align_match.group(1) if align_match else "left"

        para = document.add_paragraph()
        if align == "center":
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        for child in p.children:
            if child.name is None:
                para.add_run(str(child))
                continue

            run = para.add_run(child.get_text())

            # Text styles
            if child.name == "strong" or "font-weight:bold" in child.get("style", ""):
                run.bold = True
            if child.name == "em" or "font-style:italic" in child.get("style", ""):
                run.italic = True
            if child.name == "u" or "text-decoration:underline" in child.get("style", ""):
                run.underline = True

            # Text color
            style = child.get("style", "")
            color_match = re.search(r"color\s*:\s*(#[0-9A-Fa-f]{6})", style)
            if color_match:
                hex_color = color_match.group(1)
                run.font.color.rgb = hex_to_rgb(hex_color)

    # Save to BytesIO
    docx_io = BytesIO()
    document.save(docx_io)
    docx_io.seek(0)

    return send_file(
        docx_io,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name="converted.docx"
    )


if __name__ == '__main__':
    app.run()
