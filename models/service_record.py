from datetime import datetime, timezone
from database import db


class ServiceRecord(db.Model):
    __tablename__ = 'service_records'

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    date_from = db.Column(db.Text)
    date_to = db.Column(db.Text)
    designation = db.Column(db.Text)
    status = db.Column(db.Text)
    monthly_salary = db.Column(db.Float)
    annual_salary = db.Column(db.Float)
    station_place_of = db.Column(db.Text)
    branch = db.Column(db.Text)
    lv_abs_wo_pay = db.Column(db.Text)
    separation_date = db.Column(db.Text)
    separation_cause = db.Column(db.Text)
    sort_order = db.Column(db.Integer, default=0)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at = db.Column(
        db.DateTime,
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )

    def __repr__(self):
        return f'<ServiceRecord {self.id} emp={self.employee_id}>'
