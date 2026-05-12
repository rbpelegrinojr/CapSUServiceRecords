import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)

SQLALCHEMY_DATABASE_URI = 'sqlite:///' + os.path.join(DATA_DIR, 'records.db')
SQLALCHEMY_TRACK_MODIFICATIONS = False
SECRET_KEY = 'capsu-service-records-secret-key-2024'
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
DOCX_TEMPLATE_DIR = os.path.join(BASE_DIR, 'docx_templates')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOCX_TEMPLATE_DIR, exist_ok=True)
