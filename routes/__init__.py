from routes.employees import employees_bp
from routes.records import records_bp
from routes.import_export import import_export_bp
from routes.print_doc import print_doc_bp
from routes.settings import settings_bp

__all__ = [
    'employees_bp',
    'records_bp',
    'import_export_bp',
    'print_doc_bp',
    'settings_bp',
]
