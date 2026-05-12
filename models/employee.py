from datetime import datetime, timezone
from database import db


class Employee(db.Model):
    __tablename__ = 'employees'

    id = db.Column(db.Integer, primary_key=True)
    surname = db.Column(db.Text, nullable=False)
    given_name = db.Column(db.Text, nullable=False)
    middle_name = db.Column(db.Text)
    maiden_name = db.Column(db.Text)
    birth_date = db.Column(db.Text)
    birth_place = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at = db.Column(
        db.DateTime,
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )

    service_records = db.relationship(
        'ServiceRecord',
        backref='employee',
        lazy=True,
        cascade='all, delete-orphan',
        order_by='ServiceRecord.sort_order',
    )

    @property
    def full_name(self):
        parts = [self.surname, self.given_name]
        if self.middle_name:
            parts.append(self.middle_name)
        return ', '.join(parts)

    def __repr__(self):
        return f'<Employee {self.surname}, {self.given_name}>'
