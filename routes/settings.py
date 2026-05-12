from flask import Blueprint, render_template, request, redirect, url_for, flash
from database import db
from models.setting import Setting

settings_bp = Blueprint('settings', __name__)


@settings_bp.route('/settings', methods=['GET', 'POST'])
def manage_settings():
    if request.method == 'POST':
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
