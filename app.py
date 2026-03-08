import os
from flask import Flask, render_template, redirect, url_for, request, flash, send_file, jsonify
from flask_login import LoginManager, login_required, current_user
from models import db, User
from auth import auth_bp
from admin import admin_bp
import io
import tempfile

# ── App factory ───────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['SECRET_KEY'] = 'rep0rt-@utomation-s3cr3t-k3y-2026'
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{os.path.join(BASE_DIR, "database.db")}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024   # 50 MB

db.init_app(app)

login_manager = LoginManager()
login_manager.login_view = 'auth.login'
login_manager.login_message = 'Please log in to access this page.'
login_manager.login_message_category = 'warning'
login_manager.init_app(app)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# Register blueprints
app.register_blueprint(auth_bp)
app.register_blueprint(admin_bp)

# ── Main routes ───────────────────────────────────────────────────────────────
@app.route('/')
@login_required
def index():
    if current_user.must_change_password:
        return redirect(url_for('auth.change_password'))
    return redirect(url_for('main_dashboard'))


@app.route('/dashboard')
@login_required
def main_dashboard():
    if current_user.must_change_password:
        return redirect(url_for('auth.change_password'))
    return render_template('dashboard.html')


@app.route('/upload', methods=['POST'])
@login_required
def upload_file():
    if current_user.must_change_password:
        return jsonify({'error': 'Please change your password first.'}), 403

    if 'file' not in request.files:
        return jsonify({'error': 'No file provided.'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'error': 'No file selected.'}), 400

    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Please upload an Excel file (.xlsx or .xls).'}), 400

    # Save uploaded file temporarily
    tmp_path = os.path.join(UPLOAD_FOLDER, f'upload_{current_user.id}.xlsx')
    file.save(tmp_path)

    try:
        from report import generate_report
        result_bytes = generate_report(tmp_path)
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"Error during report generation:\n{error_details}")
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return jsonify({
            'error': f'Processing error: {str(e)}',
            'details': error_details if app.debug else 'Enable debug mode for more details.'
        }), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    # Infer month name from filename
    original_name = file.filename.replace('.xlsx', '').replace('.xls', '')
    output_name = f'Report_{original_name}_Processed.xlsx'

    return send_file(
        io.BytesIO(result_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=output_name
    )


# ── Database seeding ──────────────────────────────────────────────────────────
def seed_admin():
    admin = User.query.filter_by(email='admin@report.com').first()
    if not admin:
        admin = User(
            email='admin@report.com',
            is_admin=True,
            must_change_password=True
        )
        admin.set_password('Metal@#357')
        db.session.add(admin)
        db.session.commit()
        print('[SEED] Admin user created: admin@report.com / Metal@#357')


if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        seed_admin()
    app.run(debug=True, host='0.0.0.0', port=5000)
