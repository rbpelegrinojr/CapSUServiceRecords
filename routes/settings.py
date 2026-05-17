import os
import tempfile
from zipfile import BadZipFile

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from flask import Blueprint, current_app, render_template, request, redirect, url_for, flash
from database import db
from models.setting import Setting

settings_bp = Blueprint('settings', __name__)


@settings_bp.route('/settings', methods=['GET', 'POST'])
def manage_settings():
    if request.method == 'POST':
        uploaded_template = request.files.get('docx_template')
        if uploaded_template and uploaded_template.filename:
            filename = uploaded_template.filename.strip()
            if not filename.lower().endswith('.docx'):
                flash('Only .docx files are allowed for the service record template.', 'danger')
                return redirect(url_for('settings.manage_settings'))

            temp_path = None
            try:
                with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
                    uploaded_template.save(temp_file.name)
                    temp_path = temp_file.name

                Document(temp_path)

                template_path = os.path.join(
                    current_app.config['DOCX_TEMPLATE_DIR'],
                    'service_record_template.docx',
                )
                os.replace(temp_path, template_path)
                temp_path = None
                flash('DOCX template uploaded successfully.', 'success')
            except (PackageNotFoundError, BadZipFile, ValueError):
                flash('Invalid DOCX template. Please upload a valid .docx file.', 'danger')
                return redirect(url_for('settings.manage_settings'))
            except OSError:
                flash('Failed to save DOCX template. Please check file permissions and try again.', 'danger')
                return redirect(url_for('settings.manage_settings'))
            except Exception:
                current_app.logger.exception('Template upload failed in settings')
                flash('Template upload failed. Please try again.', 'danger')
                return redirect(url_for('settings.manage_settings'))
            finally:
                if temp_path and os.path.exists(temp_path):
                    os.remove(temp_path)

        keys = ['certifier_name', 'certifier_title', 'university_name', 'university_location']
        for key in keys:
            val = request.form.get(key, '').strip()
            setting = Setting.query.filter_by(key=key).first()
            if setting:
                setting.value = val
            else:
                db.session.add(Setting(key=key, value=val))
        db.session.commit()
        flash('Settings updated successfully.', 'success')
        return redirect(url_for('settings.manage_settings'))

    settings = {s.key: s.value for s in Setting.query.all()}
    return render_template('settings/settings.html', settings=settings)
