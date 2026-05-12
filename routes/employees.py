from datetime import datetime, timezone

from flask import Blueprint, render_template, request, redirect, url_for, flash
from database import db
from models.employee import Employee

employees_bp = Blueprint('employees', __name__)


@employees_bp.route('/employees')
def list_employees():
    query = request.args.get('q', '').strip()
    if query:
        employees = Employee.query.filter(
            db.or_(
                Employee.surname.ilike(f'%{query}%'),
                Employee.given_name.ilike(f'%{query}%'),
            )
        ).order_by(Employee.surname).all()
    else:
        employees = Employee.query.order_by(Employee.surname).all()
    return render_template('employees/list.html', employees=employees, query=query)


@employees_bp.route('/employees/add', methods=['GET', 'POST'])
def add_employee():
    if request.method == 'POST':
        surname = request.form.get('surname', '').strip()
        given_name = request.form.get('given_name', '').strip()
        if not surname or not given_name:
            flash('Surname and Given Name are required.', 'danger')
            return render_template('employees/add.html', form=request.form)
        emp = Employee(
            surname=surname.upper(),
            given_name=given_name.upper(),
            middle_name=request.form.get('middle_name', '').strip().upper() or None,
            maiden_name=request.form.get('maiden_name', '').strip().upper() or None,
            birth_date=request.form.get('birth_date', '').strip() or None,
            birth_place=request.form.get('birth_place', '').strip() or None,
        )
        db.session.add(emp)
        db.session.commit()
        flash(f'Employee {emp.surname}, {emp.given_name} added successfully.', 'success')
        return redirect(url_for('employees.view_employee', employee_id=emp.id))
    return render_template('employees/add.html', form={})


@employees_bp.route('/employees/<int:employee_id>')
def view_employee(employee_id):
    emp = Employee.query.get_or_404(employee_id)
    return render_template('employees/view.html', employee=emp)


@employees_bp.route('/employees/<int:employee_id>/edit', methods=['GET', 'POST'])
def edit_employee(employee_id):
    emp = Employee.query.get_or_404(employee_id)
    if request.method == 'POST':
        surname = request.form.get('surname', '').strip()
        given_name = request.form.get('given_name', '').strip()
        if not surname or not given_name:
            flash('Surname and Given Name are required.', 'danger')
            return render_template('employees/edit.html', employee=emp)
        emp.surname = surname.upper()
        emp.given_name = given_name.upper()
        emp.middle_name = request.form.get('middle_name', '').strip().upper() or None
        emp.maiden_name = request.form.get('maiden_name', '').strip().upper() or None
        emp.birth_date = request.form.get('birth_date', '').strip() or None
        emp.birth_place = request.form.get('birth_place', '').strip() or None
        emp.updated_at = datetime.now(timezone.utc)
        db.session.commit()
        flash('Employee updated successfully.', 'success')
        return redirect(url_for('employees.view_employee', employee_id=emp.id))
    return render_template('employees/edit.html', employee=emp)


@employees_bp.route('/employees/<int:employee_id>/delete', methods=['POST'])
def delete_employee(employee_id):
    emp = Employee.query.get_or_404(employee_id)
    name = f'{emp.surname}, {emp.given_name}'
    db.session.delete(emp)
    db.session.commit()
    flash(f'Employee {name} and all their records have been deleted.', 'success')
    return redirect(url_for('employees.list_employees'))
