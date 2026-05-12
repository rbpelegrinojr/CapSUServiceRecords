from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


def init_db(app):
    db.init_app(app)
    with app.app_context():
        db.create_all()
        _seed_settings()


def _seed_settings():
    from models.setting import Setting
    defaults = [
        ('certifier_name', 'PET ROANA B. BATACANDOLO, DPA'),
        ('certifier_title', 'Designated HRMP'),
        ('university_name', 'CAPIZ STATE UNIVERSITY'),
        ('university_location', 'Pontevedra, Capiz'),
    ]
    for key, value in defaults:
        if not Setting.query.filter_by(key=key).first():
            db.session.add(Setting(key=key, value=value))
    db.session.commit()
