from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Education(db.Model):
    __tablename__ = 'education'
    id = db.Column(db.Integer, primary_key=True)
    school = db.Column(db.String(128), nullable=False)
    degree = db.Column(db.String(128), nullable=False)
    location = db.Column(db.String(128))
    date = db.Column(db.String(64))


class Company(db.Model):
    __tablename__ = 'companies'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(128), nullable=False)
    location = db.Column(db.String(128))


class Job(db.Model):
    __tablename__ = 'jobs'
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(128), nullable=False)
    company_id = db.Column(db.Integer, db.ForeignKey(
        'companies.id'), nullable=False)

    company = db.relationship('Company', backref='jobs')
    accomplishments = db.relationship(
        'Accomplishment', backref='job', lazy=True)


class Accomplishment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    job_id = db.Column(db.Integer, db.ForeignKey('jobs.id'), nullable=False)
    content = db.Column(db.Text, nullable=False)


class TechnicalSkill(db.Model):
    __tablename__ = 'technical_skills'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)


class Project(db.Model):
    __tablename__ = 'projects'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(128), nullable=False)
    description = db.Column(db.Text)
    link = db.Column(db.String(500))
