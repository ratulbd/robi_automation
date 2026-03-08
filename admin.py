from flask import Blueprint, render_template, redirect, url_for, request, flash
from flask_login import login_required, current_user
from functools import wraps
from models import db, User

admin_bp = Blueprint('admin', __name__)

DEFAULT_PASSWORD = 'Metal@#357'

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_admin:
            flash('Admin access required.', 'danger')
            return redirect(url_for('main_dashboard'))
        return f(*args, **kwargs)
    return decorated


@admin_bp.route('/admin')
@login_required
@admin_required
def admin_panel():
    users = User.query.order_by(User.email).all()
    return render_template('admin.html', users=users)


@admin_bp.route('/admin/add-user', methods=['POST'])
@login_required
@admin_required
def add_user():
    email = request.form.get('email', '').strip().lower()
    is_admin = request.form.get('is_admin') == 'on'

    if not email:
        flash('Email is required.', 'danger')
        return redirect(url_for('admin.admin_panel'))

    existing = User.query.filter_by(email=email).first()
    if existing:
        flash(f'User {email} already exists.', 'danger')
        return redirect(url_for('admin.admin_panel'))

    new_user = User(email=email, is_admin=is_admin, must_change_password=True)
    new_user.set_password(DEFAULT_PASSWORD)
    db.session.add(new_user)
    db.session.commit()
    flash(f'User {email} added. Default password: {DEFAULT_PASSWORD}', 'success')
    return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/remove-user/<int:user_id>', methods=['POST'])
@login_required
@admin_required
def remove_user(user_id):
    user = User.query.get_or_404(user_id)
    if user.id == current_user.id:
        flash('You cannot remove your own account.', 'danger')
        return redirect(url_for('admin.admin_panel'))
    email = user.email
    db.session.delete(user)
    db.session.commit()
    flash(f'User {email} removed.', 'success')
    return redirect(url_for('admin.admin_panel'))


@admin_bp.route('/admin/reset-password/<int:user_id>', methods=['POST'])
@login_required
@admin_required
def reset_password(user_id):
    user = User.query.get_or_404(user_id)
    user.set_password(DEFAULT_PASSWORD)
    user.must_change_password = True
    db.session.commit()
    flash(f'Password for {user.email} reset to default. They must change it on next login.', 'success')
    return redirect(url_for('admin.admin_panel'))
