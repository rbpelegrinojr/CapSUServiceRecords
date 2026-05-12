import os
import threading
import webbrowser

from flask import Flask, render_template

import config
from database import db, init_db
from models.employee import Employee
from models.service_record import ServiceRecord
from routes.employees import employees_bp
from routes.records import records_bp
from routes.import_export import import_export_bp
from routes.print_doc import print_doc_bp
from routes.settings import settings_bp


def create_app():
    app = Flask(__name__)
    app.config.from_object(config)

    db.init_app(app)

    app.register_blueprint(employees_bp)
    app.register_blueprint(records_bp)
    app.register_blueprint(import_export_bp)
    app.register_blueprint(print_doc_bp)
    app.register_blueprint(settings_bp)

    with app.app_context():
        db.create_all()
        _seed_settings()
        _ensure_docx_template()
        _ensure_sample_data()

    @app.route('/')
    def index():
        total_employees = Employee.query.count()
        total_records = ServiceRecord.query.count()
        recent_employees = Employee.query.order_by(Employee.created_at.desc()).limit(5).all()
        return render_template(
            'index.html',
            total_employees=total_employees,
            total_records=total_records,
            recent_employees=recent_employees,
        )

    return app


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


def _ensure_docx_template():
    template_path = os.path.join(config.DOCX_TEMPLATE_DIR, 'service_record_template.docx')
    if not os.path.exists(template_path):
        try:
            import create_template
            create_template.create_template(template_path)
        except Exception as e:
            print(f'Warning: Could not create DOCX template: {e}')


def _ensure_sample_data():
    sample_path = os.path.join(
        os.path.dirname(__file__), 'sample_data', 'sample_service_records.xlsx'
    )
    if not os.path.exists(sample_path):
        try:
            import create_sample_data
            create_sample_data.create_sample(sample_path)
        except Exception as e:
            print(f'Warning: Could not create sample data: {e}')


def open_browser():
    webbrowser.open('http://127.0.0.1:5000')


if __name__ == '__main__':
    app = create_app()
    threading.Timer(1.5, open_browser).start()
    app.run(port=5000, debug=False, use_reloader=False)
