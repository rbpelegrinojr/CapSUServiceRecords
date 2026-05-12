from datetime import datetime, timezone

from flask import Blueprint, render_template, request, redirect, url_for, flash
from database import db
from models.employee import Employee
from models.service_record import ServiceRecord

records_bp = Blueprint('records', __name__)


@records_bp.route('/records/add/<int:employee_id>', methods=['GET', 'POST'])
def add_record(employee_id):
    emp = Employee.query.get_or_404(employee_id)
    if request.method == 'POST':
        try:
            monthly = request.form.get('monthly_salary', '').strip()
            annual = request.form.get('annual_salary', '').strip()
            max_sort = db.session.query(
                db.func.max(ServiceRecord.sort_order)
            ).filter_by(employee_id=employee_id).scalar() or 0
            rec = ServiceRecord(
                employee_id=employee_id,
                date_from=request.form.get('date_from', '').strip() or None,
                date_to=request.form.get('date_to', '').strip() or None,
                designation=request.form.get('designation', '').strip() or None,
                status=request.form.get('status', '').strip() or None,
                monthly_salary=float(monthly) if monthly else None,
                annual_salary=float(annual) if annual else None,
                station_place_of=request.form.get('station_place_of', '').strip() or None,
                branch=request.form.get('branch', '').strip() or None,
                lv_abs_wo_pay=request.form.get('lv_abs_wo_pay', '').strip() or None,
                separation_date=request.form.get('separation_date', '').strip() or None,
                separation_cause=request.form.get('separation_cause', '').strip() or None,
                sort_order=max_sort + 1,
            )
            db.session.add(rec)
            db.session.commit()
            flash('Service record added successfully.', 'success')
            return redirect(url_for('employees.view_employee', employee_id=employee_id))
        except ValueError:
            flash('Invalid salary value. Please enter a valid number.', 'danger')
    return render_template('records/add.html', employee=emp, form=request.form)


@records_bp.route('/records/<int:record_id>/edit', methods=['GET', 'POST'])
def edit_record(record_id):
    rec = ServiceRecord.query.get_or_404(record_id)
    emp = rec.employee
    if request.method == 'POST':
        try:
            monthly = request.form.get('monthly_salary', '').strip()
            annual = request.form.get('annual_salary', '').strip()
            rec.date_from = request.form.get('date_from', '').strip() or None
            rec.date_to = request.form.get('date_to', '').strip() or None
            rec.designation = request.form.get('designation', '').strip() or None
            rec.status = request.form.get('status', '').strip() or None
            rec.monthly_salary = float(monthly) if monthly else None
            rec.annual_salary = float(annual) if annual else None
            rec.station_place_of = request.form.get('station_place_of', '').strip() or None
            rec.branch = request.form.get('branch', '').strip() or None
            rec.lv_abs_wo_pay = request.form.get('lv_abs_wo_pay', '').strip() or None
            rec.separation_date = request.form.get('separation_date', '').strip() or None
            rec.separation_cause = request.form.get('separation_cause', '').strip() or None
            rec.updated_at = datetime.now(timezone.utc)
            db.session.commit()
            flash('Service record updated successfully.', 'success')
            return redirect(url_for('employees.view_employee', employee_id=emp.id))
        except ValueError:
            flash('Invalid salary value. Please enter a valid number.', 'danger')
    return render_template('records/edit.html', record=rec, employee=emp)


@records_bp.route('/records/<int:record_id>/delete', methods=['POST'])
def delete_record(record_id):
    rec = ServiceRecord.query.get_or_404(record_id)
    employee_id = rec.employee_id
    db.session.delete(rec)
    db.session.commit()
    flash('Service record deleted successfully.', 'success')
    return redirect(url_for('employees.view_employee', employee_id=employee_id))
