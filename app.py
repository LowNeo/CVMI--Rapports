"""
Application web de génération de rapports — L'OPTINÉO / CVMI
Accès privé par login/mot de passe
"""

import os, uuid, tempfile
from flask import (Flask, render_template, request, redirect,
                   url_for, session, send_file, flash)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from functools import wraps

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'changez-cette-cle-en-prod-2024')

# ─── UTILISATEURS ────────────────────────────────────────
# Définis dans les variables d'environnement Render
# Format : USER_PASSWORD et ADMIN_PASSWORD
USERS = {
    'client': {
        'password_hash': generate_password_hash(
            os.environ.get('USER_PASSWORD', 'cvmi2024')
        ),
        'role': 'client',
        'display': 'Client CVMI',
    },
    'admin': {
        'password_hash': generate_password_hash(
            os.environ.get('ADMIN_PASSWORD', 'admin-loptineo-2024')
        ),
        'role': 'admin',
        'display': 'Administrateur',
    },
}

ALLOWED_EXT = {'xlsx', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT

# ─── AUTH ────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'username' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password', '')
        user = USERS.get(username)
        if user and check_password_hash(user['password_hash'], password):
            session['username'] = username
            session['role']     = user['role']
            session['display']  = user['display']
            return redirect(url_for('index'))
        flash('Identifiant ou mot de passe incorrect.', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ─── PAGES ───────────────────────────────────────────────

@app.route('/')
@login_required
def index():
    return render_template('index.html')

@app.route('/rapport-pei', methods=['GET', 'POST'])
@login_required
def rapport_pei():
    if request.method == 'POST':
        if 'file' not in request.files or request.files['file'].filename == '':
            flash("Veuillez sélectionner un fichier.", 'error')
            return redirect(request.url)

        f = request.files['file']
        if not allowed_file(f.filename):
            flash("Format non supporté. Utilisez un fichier .xlsx ou .csv", 'error')
            return redirect(request.url)

        try:
            # Sauvegarde temporaire du fichier uploadé
            suffix = '.' + f.filename.rsplit('.', 1)[1].lower()
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_in:
                f.save(tmp_in.name)
                source_path = tmp_in.name

            # Fichier de sortie temporaire
            out_path = tempfile.mktemp(suffix='.xlsx')

            # Génération du rapport
            from generate_rapport_pei import build_rapport_pei
            build_rapport_pei(source_path, out_path)

            os.unlink(source_path)

            return send_file(
                out_path,
                as_attachment=True,
                download_name='rapport_pei.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        except Exception as e:
            flash(f"Erreur lors de la génération : {str(e)}", 'error')
            return redirect(request.url)

    return render_template('rapport_pei.html')

@app.route('/rapport-rps', methods=['GET', 'POST'])
@login_required
def rapport_rps():
    if request.method == 'POST':
        if 'file' not in request.files or request.files['file'].filename == '':
            flash("Veuillez sélectionner un fichier.", 'error')
            return redirect(request.url)

        f = request.files['file']
        entreprise = request.form.get('entreprise', 'Entreprise').strip()

        if not allowed_file(f.filename):
            flash("Format non supporté. Utilisez un fichier .xlsx ou .csv", 'error')
            return redirect(request.url)

        try:
            suffix = '.' + f.filename.rsplit('.', 1)[1].lower()
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_in:
                f.save(tmp_in.name)
                source_path = tmp_in.name

            out_path = tempfile.mktemp(suffix='.xlsx')

            from generate_rapport_rps import build_excel
            build_excel(source_path, out_path, entreprise)

            os.unlink(source_path)

            safe_name = secure_filename(entreprise) or 'entreprise'
            return send_file(
                out_path,
                as_attachment=True,
                download_name=f'rapport_rps_{safe_name}.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        except Exception as e:
            flash(f"Erreur lors de la génération : {str(e)}", 'error')
            return redirect(request.url)

    return render_template('rapport_rps.html')

# ─── PAGE ADMIN ──────────────────────────────────────────

@app.route('/admin')
@login_required
def admin():
    if session.get('role') != 'admin':
        flash("Accès réservé à l'administrateur.", 'error')
        return redirect(url_for('index'))
    return render_template('admin.html', users=USERS)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
