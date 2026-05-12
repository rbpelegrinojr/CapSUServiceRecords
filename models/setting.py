from database import db


class Setting(db.Model):
    __tablename__ = 'settings'

    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.Text, unique=True, nullable=False)
    value = db.Column(db.Text)

    def __repr__(self):
        return f'<Setting {self.key}={self.value}>'
