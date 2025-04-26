from flask import Flask, jsonify, render_template, send_file
from config import SQLALCHEMY_DATABASE_URI
from models import db, Experience, Accomplishment
from docx.shared import Pt, RGBColor
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = SQLALCHEMY_DATABASE_URI
db.init_app(app)


def extract_keywords(text):
    stopwords = {'and', 'the', 'in', 'on', 'with', 'an', 'a',
                 'to', 'for', 'we', 'are', 'of', 'is', 'be', 'as'}
    words = text.lower().split()
    keywords = [word.strip(".,")
                for word in words if word not in stopwords and len(word) > 2]
    return set(keywords)


def add_hyperlink(paragraph, url, text):
    """
    Add a hyperlink to a paragraph that looks like a real hyperlink (blue and underlined).
    """
    # Create the w:hyperlink tag and add needed values
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    # Create a w:r element
    new_run = OxmlElement('w:r')

    # Create a w:rPr element (run properties)
    rPr = OxmlElement('w:rPr')

    # Set color to blue
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0000FF')  # Blue color
    rPr.append(color)

    # Set underline
    underline = OxmlElement('w:u')
    underline.set(qn('w:val'), 'single')
    rPr.append(underline)

    # Style the run like a hyperlink
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'Hyperlink')
    rPr.append(rStyle)

    new_run.append(rPr)

    # Create a w:t element (the text inside the hyperlink)
    text_elem = OxmlElement('w:t')
    text_elem.text = text
    new_run.append(text_elem)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_tab_stop(paragraph, position_twips):
    """Add tab stop at specific position (measured in twips)."""
    pPr = paragraph._p.get_or_add_pPr()
    tabs = pPr.find(qn('w:tabs'))
    if tabs is None:
        tabs = OxmlElement('w:tabs')
        pPr.append(tabs)
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), str(position_twips))
    tabs.append(tab)


@app.route('/experiences', methods=['GET'])
def get_experiences():
    experiences = Experience.query.all()
    data = []
    for exp in experiences:
        data.append({
            'job_title': exp.job_title,
            'company': exp.company,
            'summary': exp.summary,
            'accomplishments': [a.content for a in exp.accomplishments]
        })
    return jsonify(data)


@app.route('/match', methods=['GET'])
def match_job_description():
    job_description = """
    We are seeking an experienced Automation Developer skilled in UiPath and RPA systems.
    The role involves building bots, optimizing workflows, and automating manual tasks.
    """
    keywords = extract_keywords(job_description)
    print("Extracted keywords:", keywords)

    matches = []
    all_accomplishments = Accomplishment.query.all()

    for acc in all_accomplishments:
        if any(keyword.lower() in acc.content.lower() for keyword in keywords):
            matches.append(acc.content)

    return jsonify({
        'job_description': job_description.strip(),
        'matched_bullets': matches
    })


@app.route('/resume', methods=['GET'])
def generate_resume():
    job_description = """
    Seeking an RPA Developer with expertise in UiPath, automation, and strong problem-solving skills.
    """
    keywords = extract_keywords(job_description)

    matched = []
    for acc in Accomplishment.query.all():
        if any(keyword.lower() in acc.content.lower() for keyword in keywords):
            matched.append(acc.content)

    # Create the Word document
    doc = Document()

    # 1. Header
    doc.add_heading("Christopher A. Roberts", level=0)

    header_para = doc.add_paragraph()
    header_para.paragraph_format.space_after = Pt(0)
    header_para.paragraph_format.space_before = Pt(0)

    header_para.add_run("Philadelphia, PA | ").font.name = 'Calibri'
    header_para.add_run(
        "Christopher.roberts11220@gmail.com | ").font.name = 'Calibri'
    header_para.add_run("908-963-0613 | ").font.name = 'Calibri'

    # Add GitHub real link
    add_hyperlink(header_para, "https://github.com/Chris1112220", "GitHub")

    header_para.add_run(" | ").font.name = 'Calibri'

    # Add LinkedIn real link
    add_hyperlink(
        header_para, "https://www.linkedin.com/in/christopher-roberts-philadelphia/", "LinkedIn")

    # Technical Skills
    doc.add_heading("Technical Skills", level=1)
    doc.add_paragraph(
        "RPA Tools: UiPath Studio, StudioX, Orchestrator\n"
        "Languages: Python, Java, C, HTML, CSS\n"
        "Databases: MySQL, PostgreSQL, SQLite\n"
        "Tools: GitHub, SAP, Salesforce, QuickBooks, Excel\n"
        "Other: Process Mapping, Agile Project Delivery"
    )
    doc.add_paragraph()

    # Professional Experience
    doc.add_heading("Professional Experience", level=1)
    for bullet in matched:
        doc.add_paragraph(bullet, style='List Bullet')
    doc.add_paragraph()

    # Projects
    doc.add_heading("Projects", level=1)
    doc.add_paragraph(
        "• Journal Entry Bot – Automated journal entry processing with UiPath.")
    doc.add_paragraph(
        "• Auto Bank Reconciliation Bot – Automated reconciliation of 50+ bank accounts.")
    doc.add_paragraph(
        "• ACH Download Bot – Automated ACH statement downloads using UiPath and Excel.")
    doc.add_paragraph()

    # === EDUCATION ===
    doc.add_heading("Education", level=1)

    # Drexel + Temple in one "invisible" structure, using only line breaks
    edu_para = doc.add_paragraph()
    edu_para.paragraph_format.space_after = Pt(0)
    edu_para.paragraph_format.space_before = Pt(0)
    add_tab_stop(edu_para, 9360)

    # Drexel
    drexel_run1 = edu_para.add_run("Drexel University")
    drexel_run1.bold = True
    drexel_run1.font.name = 'Calibri'
    drexel_run1.font.size = Pt(11)
    edu_para.add_run(
        ", College of Computing and Informatics, Philadelphia, PA\t")
    drexel_run2 = edu_para.add_run("January 2024")
    drexel_run2.font.name = 'Calibri'
    drexel_run2.font.size = Pt(11)

    edu_para.add_run("\n")  # Just one simple line break

    drexel_degree = edu_para.add_run(
        "Post-Baccalaureate Graduate Certificate in Computer Science Foundations")
    drexel_degree.font.name = 'Calibri'
    drexel_degree.font.size = Pt(11)
    edu_para.add_run("\n")
    # Just one simple line break (between Drexel and Temple)
    edu_para.add_run("\n")

    # Temple
    temple_run1 = edu_para.add_run("Temple University")
    temple_run1.bold = True
    temple_run1.font.name = 'Calibri'
    temple_run1.font.size = Pt(11)
    edu_para.add_run(", Fox School of Business — Philadelphia, PA\t")
    temple_run2 = edu_para.add_run("January 2012")
    temple_run2.font.name = 'Calibri'
    temple_run2.font.size = Pt(11)

    edu_para.add_run("\n")

    temple_degree = edu_para.add_run("BBA, Finance")
    temple_degree.font.name = 'Calibri'
    temple_degree.font.size = Pt(11)

    # Save the document
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Christopher_Roberts_Resume.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == '__main__':
    app.run(debug=True)
