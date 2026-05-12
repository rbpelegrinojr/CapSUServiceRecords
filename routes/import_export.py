import io
import os

import pandas as pd
from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from werkzeug.utils import secure_filename
from database import db
from models.employee import Employee
from models.service_record import ServiceRecord

import_export_bp = Blueprint('import_export', __name__)

ALLOWED_EXTENSIONS = {'xlsx'}


def _allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@import_export_bp.route('/import', methods=['GET', 'POST'])
def import_records():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part in the request.', 'danger')
            return redirect(url_for('import_export.import_records'))
        f = request.files['file']
        if f.filename == '':
            flash('No file selected.', 'danger')
            return redirect(url_for('import_export.import_records'))
        if not _allowed_file(f.filename):
            flash('Only .xlsx files are allowed.', 'danger')
            return redirect(url_for('import_export.import_records'))

        try:
            df = pd.read_excel(io.BytesIO(f.read()), dtype=str)
            df.columns = [c.strip() for c in df.columns]
            df = df.where(pd.notnull(df), None)

            required_cols = {'Surname', 'Given Name'}
            missing = required_cols - set(df.columns)
            if missing:
                flash(f'Missing required columns: {", ".join(missing)}', 'danger')
                return redirect(url_for('import_export.import_records'))

            created_employees = 0
            created_records = 0
            errors = []

            for idx, row in df.iterrows():
                try:
                    surname = (row.get('Surname') or '').strip().upper()
                    given_name = (row.get('Given Name') or '').strip().upper()
                    if not surname or not given_name:
                        errors.append(f'Row {idx + 2}: Missing Surname or Given Name.')
                        continue

                    emp = Employee.query.filter(
                        db.func.upper(Employee.surname) == surname,
                        db.func.upper(Employee.given_name) == given_name,
                    ).first()

                    if not emp:
                        emp = Employee(
                            surname=surname,
                            given_name=given_name,
                            middle_name=(row.get('Middle Name') or '').strip().upper() or None,
                            birth_date=row.get('Birth Date'),
                            birth_place=row.get('Birth Place'),
                        )
                        db.session.add(emp)
                        db.session.flush()
                        created_employees += 1

                    def _safe_float(val):
                        try:
                            return float(val) if val not in (None, '') else None
                        except (ValueError, TypeError):
                            return None

                    max_sort = db.session.query(
                        db.func.max(ServiceRecord.sort_order)
                    ).filter_by(employee_id=emp.id).scalar() or 0

                    rec = ServiceRecord(
                        employee_id=emp.id,
                        date_from=row.get('Date From'),
                        date_to=row.get('Date To'),
                        designation=row.get('Designation'),
                        status=row.get('Status'),
                        monthly_salary=_safe_float(row.get('Monthly Salary')),
                        annual_salary=_safe_float(row.get('Annual Salary')),
                        station_place_of=row.get('Station/Place of'),
                        branch=row.get('Branch'),
                        lv_abs_wo_pay=row.get('LV ABS WO Pay'),
                        separation_date=row.get('Separation Date'),
                        separation_cause=row.get('Separation Cause'),
                        sort_order=max_sort + 1,
                    )
                    db.session.add(rec)
                    created_records += 1

                except Exception as e:
                    errors.append(f'Row {idx + 2}: {str(e)}')

            db.session.commit()

            msg = f'Import complete: {created_employees} employee(s) created, {created_records} record(s) imported.'
            flash(msg, 'success')
            if errors:
                for err in errors[:10]:
                    flash(err, 'warning')
                if len(errors) > 10:
                    flash(f'... and {len(errors) - 10} more errors.', 'warning')

        except Exception as e:
            db.session.rollback()
            flash(f'Import failed: {str(e)}', 'danger')

        return redirect(url_for('import_export.import_records'))

    return render_template('import_export/import.html')


@import_export_bp.route('/export/<int:employee_id>')
def export_employee(employee_id):
    emp = Employee.query.get_or_404(employee_id)
    records = ServiceRecord.query.filter_by(employee_id=employee_id).order_by(
        ServiceRecord.sort_order
    ).all()

    rows = []
    for rec in records:
        rows.append({
            'Surname': emp.surname,
            'Given Name': emp.given_name,
            'Middle Name': emp.middle_name or '',
            'Birth Date': emp.birth_date or '',
            'Birth Place': emp.birth_place or '',
            'Date From': rec.date_from or '',
            'Date To': rec.date_to or '',
            'Designation': rec.designation or '',
            'Status': rec.status or '',
            'Monthly Salary': rec.monthly_salary or '',
            'Annual Salary': rec.annual_salary or '',
            'Station/Place of': rec.station_place_of or '',
            'Branch': rec.branch or '',
            'LV ABS WO Pay': rec.lv_abs_wo_pay or '',
            'Separation Date': rec.separation_date or '',
            'Separation Cause': rec.separation_cause or '',
        })

    if not rows:
        rows = [{'Surname': emp.surname, 'Given Name': emp.given_name}]

    df = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Service Records')
    output.seek(0)

    filename = f'service_records_{emp.surname}_{emp.given_name}.xlsx'.replace(' ', '_')
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )


@import_export_bp.route('/download-sample-template')
def download_sample_template():
    sample_path = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        'sample_data',
        'sample_service_records.xlsx',
    )
    if not os.path.exists(sample_path):
        flash('Sample template not found.', 'warning')
        return redirect(url_for('import_export.import_records'))
    return send_file(
        sample_path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='sample_service_records.xlsx',
    )
