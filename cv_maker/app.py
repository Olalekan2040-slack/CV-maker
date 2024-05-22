from flask import Flask, render_template, request, send_file
import io
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from flask import Flask, render_template, request, send_file
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re





app = Flask(__name__)

def clean_text(text):
    # Remove any non-XML-compatible characters
    return re.sub(r'[^\x20-\x7E]+', '', text)

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    data = {
        'name': clean_text(request.form['name']),
        'title': clean_text(request.form['title']),
        'contact': clean_text(request.form['contact']),
        'summary': clean_text(request.form['summary']),
        'projects': clean_text(request.form['projects']),
        'education': clean_text(request.form['education']),
        'experience': clean_text(request.form['experience']),
        'skills': clean_text(request.form['skills']),
        'languages': clean_text(request.form['languages'])
    }
    
    # Generate PDF and Word document
    pdf_buffer = generate_pdf(data)
    word_buffer = generate_word(data)
    
    return send_file(word_buffer, as_attachment=True, download_name='resume.docx')

def generate_pdf(data):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    
    # Header
    c.setFont("Helvetica-Bold", 16)
    c.drawString(30, height - 40, data['name'])
    c.setFont("Helvetica", 12)
    c.drawString(30, height - 60, data['title'])
    c.drawString(30, height - 80, data['contact'])
    
    # Summary
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, height - 120, "Summary")
    c.setFont("Helvetica", 12)
    c.drawString(30, height - 140, data['summary'])
    
    # Projects
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, height - 180, "Projects")
    c.setFont("Helvetica", 12)
    c.drawString(30, height - 200, data['projects'])
    
    # Education
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, height - 240, "Education")
    c.setFont("Helvetica", 12)
    c.drawString(30, height - 260, data['education'])
    
    # Experience
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, height - 300, "Experience")
    c.setFont("Helvetica", 12)
    c.drawString(30, height - 320, data['experience'])
    
    # Skills
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, height - 360, "Skills")
    c.setFont("Helvetica", 12)
    c.drawString(30, height - 380, data['skills'])
    
    # Languages
    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, height - 420, "Languages")
    c.setFont("Helvetica", 12)
    c.drawString(30, height - 440, data['languages'])
    
    c.save()
    
    buffer.seek(0)
    return buffer




def generate_word(data):
    doc = Document()
    
    doc.add_heading(data['name'], 0)
    doc.add_paragraph(data['title'])

    # Center-align the name, title, and contact information
    # def add_centered_paragraph(text, bold=False, font_size=26):
    #     para = doc.add_paragraph()
    #     run = para.add_run(text)
    #     if bold:
    #         run.bold = True
    #     run.font.size = Pt(font_size)
    #     para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # add_centered_paragraph(data['name'], bold=True, font_size=16)
    # add_centered_paragraph(data['title'], font_size=14)
    # add_centered_paragraph(data['contact'], font_size=12)

    # Add a line separator
    def add_line_separator():
        p = doc.add_paragraph()
        p.add_run().add_break()
        p = doc.add_paragraph()
        run = p.add_run()
        p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        border = OxmlElement('w:bottom')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '6')
        border.set(qn('w:space'), '1')
        border.set(qn('w:color'), '000000')
        p._element.get_or_add_pPr().append(border)

    add_line_separator()

    # Add a section with a heading and content
    def add_section(heading, content):
        doc.add_heading(heading, level=2)
        for item in content.split('\n'):
            if item.strip():
                doc.add_paragraph(item.strip(), style='ListBullet')
        add_line_separator()

    # Create a table for two-column layout
    table = doc.add_table(rows=1, cols=2)
    table.autofit = True
    table.columns[0].width = Pt(200)
    table.columns[1].width = Pt(200)
    
    # Add left column content
    left_cell = table.cell(0, 0)
    left_cell.text = ''
    add_section('Summary', data['summary'])
    add_section('Education', data['education'])
    add_section('Projects', data['projects'])

    # Add right column content
    right_cell = table.cell(0, 1)
    right_cell.text = ''
    add_section('Experience', data['experience'])
    add_section('Skills', data['skills'])
    add_section('Languages', data['languages'])
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer

if __name__ == '__main__':
    app.run(debug=True)
