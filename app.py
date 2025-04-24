from flask import Flask, jsonify
from config import SQLALCHEMY_DATABASE_URI
from models import db, Experience, Accomplishment
from flask import Flask, jsonify, render_template
from docx import Document
from flask import send_file
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
    We are looking for an RPA Developer with expertise in UiPath and automation of financial processes.
    """
    keywords = extract_keywords(job_description)

    matched = []
    for acc in Accomplishment.query.all():
        if any(keyword.lower() in acc.content.lower() for keyword in keywords):
            matched.append(acc.content)

    return render_template(
        'resume.html',
        name="Christopher A. Roberts",
        job_title="RPA Developer",
        bullets=matched
    )


@app.route('/resume-docx', methods=['GET'])
def generate_resume_docx():
    job_description = """
    Seeking a skilled RPA Developer with UiPath experience and a strong automation mindset.
    """
    keywords = extract_keywords(job_description)

    matched = []
    for acc in Accomplishment.query.all():
        if any(keyword.lower() in acc.content.lower() for keyword in keywords):
            matched.append(acc.content)

    # Create DOCX
    doc = Document()
    doc.add_heading("Christopher A. Roberts", level=1)
    doc.add_paragraph("Target Role: RPA Developer")

    # Education section
    doc.add_heading("Education", level=2)
    doc.add_paragraph(
        "Drexel University — Post-Baccalaureate Certificate in CS Foundations (Jan 2024)")
    doc.add_paragraph("Temple University — BBA in Finance (Jan 2012)")

    # Accomplishments section
    doc.add_heading("Relevant Accomplishments", level=2)
    for bullet in matched:
        doc.add_paragraph(bullet, style='List Bullet')

    # Export to memory
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Resume_RPA_Developer.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == '__main__':
    app.run(debug=True)
