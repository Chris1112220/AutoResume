from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Experience(db.Model):
    __tablename__ = 'experiences'
    id = db.Column(db.Integer, primary_key=True)
    job_title = db.Column(db.String)
    company = db.Column(db.String)
    start_date = db.Column(db.Date)
    end_date = db.Column(db.Date)
    summary = db.Column(db.Text)
    accomplishments = db.relationship(
        'Accomplishment', backref='experience', cascade='all, delete')


class Accomplishment(db.Model):
    __tablename__ = 'accomplishments'
    id = db.Column(db.Integer, primary_key=True)
    experience_id = db.Column(db.Integer, db.ForeignKey('experiences.id'))
    content = db.Column(db.Text)
    tags = db.Column(db.ARRAY(db.String))  # PostgreSQL array field
