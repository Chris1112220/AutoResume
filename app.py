from flask import Flask, jsonify, send_file
from config import SQLALCHEMY_DATABASE_URI
from models import db, Education, Company, Job, Accomplishment, TechnicalSkill, Project
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
    """Add clickable hyperlink looking like real hyperlink."""
    part = paragraph.part
    r_id = part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0000FF')
    rPr.append(color)
    underline = OxmlElement('w:u')
    underline.set(qn('w:val'), 'single')
    rPr.append(underline)
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'Hyperlink')
    rPr.append(rStyle)
    new_run.append(rPr)
    text_elem = OxmlElement('w:t')
    text_elem.text = text
    new_run.append(text_elem)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_tab_stop(paragraph, position_twips):
    pPr = paragraph._p.get_or_add_pPr()
    tabs = pPr.find(qn('w:tabs'))
    if tabs is None:
        tabs = OxmlElement('w:tabs')
        pPr.append(tabs)
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), str(position_twips))
    tabs.append(tab)


@app.route('/resume', methods=['GET'])
def generate_resume():
    job_description = """
    Seeking an RPA Developer with expertise in UiPath, automation, and strong problem-solving skills.
    """
    keywords = extract_keywords(job_description)

    matched_accomplishments = []
    all_accomplishments = Accomplishment.query.all()

    for acc in all_accomplishments:
        if any(keyword.lower() in acc.description.lower() for keyword in keywords):
            matched_accomplishments.append(acc)

    # Fetch technical skills
    tech_skills = TechnicalSkill.query.all()

    # Fetch projects
    projects = Project.query.all()

    # Fetch jobs with companies
    jobs = db.session.query(Job, Company).join(
        Company, Job.company_id == Company.id).all()

    # Start Word document
    doc = Document()

    # HEADER
    doc.add_heading("Christopher A. Roberts", level=0)
    header_para = doc.add_paragraph()
    header_para.paragraph_format.space_after = Pt(0)
    header_para.paragraph_format.space_before = Pt(0)

    header_para.add_run("Philadelphia, PA | ").font.name = 'Calibri'
    header_para.add_run(
        "Christopher.roberts11220@gmail.com | ").font.name = 'Calibri'
    header_para.add_run("908-963-0613 | ").font.name = 'Calibri'
    add_hyperlink(header_para, "https://github.com/Chris1112220", "GitHub")
    header_para.add_run(" | ").font.name = 'Calibri'
    add_hyperlink(
        header_para, "https://www.linkedin.com/in/christopher-roberts-philadelphia/", "LinkedIn")

    doc.add_paragraph()

    # TECHNICAL SKILLS
    doc.add_heading("Technical Skills", level=1)
    for skill in tech_skills:
        doc.add_paragraph(skill.name, style='List Bullet')
    doc.add_paragraph()

    # PROFESSIONAL EXPERIENCE
    doc.add_heading("Professional Experience", level=1)
    for job, company in jobs:
        job_header = doc.add_paragraph()
        job_header.add_run(f"{job.title} | {company.name}").bold = True
        job_header.paragraph_format.space_after = Pt(0)

        # Match bullets for this job
        for acc in matched_accomplishments:
            if acc.job_id == job.id:
                doc.add_paragraph(acc.description, style='List Bullet')

        doc.add_paragraph()

    # PROJECTS
    doc.add_heading("Projects", level=1)
    for proj in projects:
        project_para = doc.add_paragraph()
        project_para.add_run(f"{proj.name} – ").bold = True
        project_para.add_run(proj.description)
        if proj.link:
            project_para.add_run("\n")
            add_hyperlink(project_para, proj.link, proj.link)

    doc.add_paragraph()

    # EDUCATION (keeping same format you wanted perfect)
    doc.add_heading("Education", level=1)

    edu_para = doc.add_paragraph()
    edu_para.paragraph_format.space_after = Pt(0)
    edu_para.paragraph_format.space_before = Pt(0)
    add_tab_stop(edu_para, 9360)

    drexel_run1 = edu_para.add_run("Drexel University")
    drexel_run1.bold = True
    drexel_run1.font.name = 'Calibri'
    drexel_run1.font.size = Pt(11)
    edu_para.add_run(
        ", College of Computing and Informatics, Philadelphia, PA\t")
    drexel_run2 = edu_para.add_run("January 2024")
    drexel_run2.font.name = 'Calibri'
    drexel_run2.font.size = Pt(11)

    edu_para.add_run("\n")
    drexel_degree = edu_para.add_run(
        "Post-Baccalaureate Graduate Certificate in Computer Science Foundations")
    drexel_degree.font.name = 'Calibri'
    drexel_degree.font.size = Pt(11)

    edu_para.add_run("\n\n")

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

    # SAVE
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
