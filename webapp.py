#!/usr/bin/env python3
"""
Webapp Sophie — Génération de vidéos L'App de Sophie.

Sophie upload ses fichiers Excel, choisit la semaine,
et récupère la vidéo MP4 générée automatiquement.

Usage:
    python webapp.py
    → Ouvrir http://localhost:8080
"""

import datetime
import functools
import os
import shutil
import tempfile
import uuid

from dotenv import load_dotenv

load_dotenv()

from flask import (
    Flask, render_template_string, request, send_file,
    redirect, url_for, flash, session,
)
from werkzeug.utils import secure_filename

from src.data_loader import (
    load_projects,
    filter_events_this_week,
    filter_recos_this_week,
    format_event_card,
    format_reco_card,
    get_week_number,
    get_week_bounds,
)
from src.logger import get_logger
from src.slide_renderer import build_slide_sequence, render_slides_to_images
from src.storage import upload_video, get_video_url
from src.supabase_client import get_supabase
from src.video_builder import build_video

log = get_logger()


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24))

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(PROJECT_DIR, "uploads")
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def _supabase_configured() -> bool:
    """Vérifie si Supabase est configuré."""
    return get_supabase() is not None


def login_required(f):
    """Décorateur : redirige vers /login si pas de session active."""
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        if not _supabase_configured():
            # Pas de Supabase → pas d'auth, accès libre
            return f(*args, **kwargs)
        if not session.get("user"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


# ---------------------------------------------------------------------------
# CSS commun (réutilisé par toutes les pages)
# ---------------------------------------------------------------------------

COMMON_CSS = """
    :root {
        --black: #000000;
        --white: #FFFFFF;
        --gray-dark: #4f4e4d;
        --gray-medium: #c1c0bc;
        --gray-light: #f6f5f4;
    }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
        font-family: 'Aptos', system-ui, 'Segoe UI', sans-serif;
        background: var(--gray-light);
        color: var(--gray-dark);
        min-height: 100vh;
        display: flex;
        flex-direction: column;
        align-items: center;
    }
    .header {
        width: 100%;
        background: var(--black);
        color: var(--white);
        padding: 32px 0;
        text-align: center;
    }
    .header h1 {
        font-size: 28px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 2px;
    }
    .header p {
        font-size: 14px;
        color: var(--gray-medium);
        margin-top: 8px;
    }
    .container {
        max-width: 640px;
        width: 100%;
        padding: 40px 24px;
    }
    .card {
        background: var(--white);
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        padding: 32px;
        margin-bottom: 24px;
    }
    .card h2 {
        font-size: 18px;
        font-weight: 700;
        text-transform: uppercase;
        color: var(--black);
        margin-bottom: 8px;
    }
    .card .separator {
        width: 40px;
        height: 3px;
        background: var(--black);
        margin-bottom: 20px;
    }
    label {
        display: block;
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        color: var(--gray-medium);
        letter-spacing: 0.5px;
        margin-bottom: 6px;
        margin-top: 16px;
    }
    label:first-of-type { margin-top: 0; }
    input[type="file"], input[type="date"], input[type="email"], input[type="password"] {
        width: 100%;
        padding: 12px;
        border: 1px solid var(--gray-medium);
        border-radius: 4px;
        font-size: 15px;
        font-family: inherit;
        background: var(--gray-light);
    }
    input[type="file"] { cursor: pointer; }
    button {
        display: block;
        width: 100%;
        margin-top: 24px;
        padding: 14px;
        background: var(--black);
        color: var(--white);
        border: none;
        font-size: 16px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        cursor: pointer;
        border-radius: 4px;
        transition: opacity 0.2s;
    }
    button:hover { opacity: 0.85; }
    button:disabled {
        opacity: 0.4;
        cursor: wait;
    }
    .flash {
        padding: 12px 16px;
        border-radius: 4px;
        margin-bottom: 16px;
        font-size: 14px;
    }
    .flash.error {
        background: #fdecea;
        color: #b71c1c;
        border: 1px solid #ef9a9a;
    }
    .flash.success {
        background: #e8f5e9;
        color: #1b5e20;
        border: 1px solid #a5d6a7;
    }
    .result {
        text-align: center;
        padding: 24px;
    }
    .result a {
        display: inline-block;
        padding: 14px 32px;
        background: var(--black);
        color: var(--white);
        text-decoration: none;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        border-radius: 4px;
    }
    .result a:hover { opacity: 0.85; }
    .result .info {
        margin-top: 12px;
        font-size: 13px;
        color: var(--gray-medium);
    }
    .footer {
        text-align: center;
        padding: 20px;
        font-size: 11px;
        color: var(--gray-medium);
    }
    .spinner {
        display: none;
        margin: 16px auto 0;
        width: 32px;
        height: 32px;
        border: 3px solid var(--gray-medium);
        border-top: 3px solid var(--black);
        border-radius: 50%;
        animation: spin 0.8s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .nav-bar {
        width: 100%;
        background: var(--black);
        display: flex;
        justify-content: flex-end;
        padding: 8px 24px;
    }
    .nav-bar a {
        color: var(--gray-medium);
        text-decoration: none;
        font-size: 13px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .nav-bar a:hover { color: var(--white); }
"""


# ---------------------------------------------------------------------------
# Templates
# ---------------------------------------------------------------------------

LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sophie — Connexion</title>
    <style>""" + COMMON_CSS + """</style>
</head>
<body>
    <div class="header">
        <h1>L'App de Sophie</h1>
        <p>WMH Project — Connexion</p>
    </div>
    <div class="container">
        {% with messages = get_flashed_messages(with_categories=true) %}
        {% for category, message in messages %}
        <div class="flash {{ category }}">{{ message }}</div>
        {% endfor %}
        {% endwith %}

        <div class="card">
            <h2>Connexion</h2>
            <div class="separator"></div>
            <form method="POST">
                <label for="email">Email</label>
                <input type="email" name="email" id="email" required autocomplete="email">

                <label for="password">Mot de passe</label>
                <input type="password" name="password" id="password" required autocomplete="current-password">

                <button type="submit">Se connecter</button>
            </form>
        </div>
    </div>
    <div class="footer">&copy;WMH Project — {{ year }}</div>
</body>
</html>
"""

MAIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sophie — L'App de Sophie</title>
    <style>""" + COMMON_CSS + """</style>
</head>
<body>
    <div class="header">
        <h1>L'App de Sophie</h1>
        <p>WMH Project — Génération automatique de vidéos</p>
    </div>
    {% if user_email %}
    <div class="nav-bar">
        <a href="{{ url_for('logout') }}">Déconnexion ({{ user_email }})</a>
    </div>
    {% endif %}
    <div class="container">
        {% with messages = get_flashed_messages(with_categories=true) %}
        {% for category, message in messages %}
        <div class="flash {{ category }}">{{ message }}</div>
        {% endfor %}
        {% endwith %}

        {% if video_url %}
        <div class="card">
            <h2>Vidéo prête</h2>
            <div class="separator"></div>
            <div class="result">
                <a href="{{ video_url }}">Télécharger le MP4</a>
                <div class="info">{{ video_info }}</div>
            </div>
        </div>
        {% endif %}

        <div class="card">
            <h2>Générer une vidéo</h2>
            <div class="separator"></div>
            <form method="POST" enctype="multipart/form-data" id="genForm">
                <label for="file_gagne">Fichier « Gagné » (.xlsx) — événements qui jouent</label>
                <input type="file" name="file_gagne" id="file_gagne" accept=".xlsx" required>

                <label for="file_pipe">Fichier « Pipe » (.xlsx) — recos à rendre</label>
                <input type="file" name="file_pipe" id="file_pipe" accept=".xlsx" required>

                <label for="ref_date">Semaine de référence</label>
                <input type="date" name="ref_date" id="ref_date" value="{{ today }}">

                <button type="submit" id="submitBtn">Générer la vidéo</button>
                <div class="spinner" id="spinner"></div>
            </form>
        </div>
    </div>
    <div class="footer">&copy;WMH Project — {{ year }}</div>

    <script>
        document.getElementById('genForm').addEventListener('submit', function() {
            document.getElementById('submitBtn').disabled = true;
            document.getElementById('submitBtn').textContent = 'Génération en cours...';
            document.getElementById('spinner').style.display = 'block';
        });
    </script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Routes — Auth
# ---------------------------------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    if not _supabase_configured():
        return redirect(url_for("index"))

    if session.get("user"):
        return redirect(url_for("index"))

    if request.method == "POST":
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "")

        if not email or not password:
            flash("Email et mot de passe requis.", "error")
            return redirect(url_for("login"))

        sb = get_supabase()
        try:
            response = sb.auth.sign_in_with_password({
                "email": email,
                "password": password,
            })
            session["user"] = {
                "email": response.user.email,
                "access_token": response.session.access_token,
            }
            log.info(f"Connexion réussie : {email}")
            return redirect(url_for("index"))
        except Exception as e:
            log.warning(f"Échec de connexion pour {email} : {e}")
            flash("Email ou mot de passe incorrect.", "error")
            return redirect(url_for("login"))

    return render_template_string(
        LOGIN_TEMPLATE,
        year=datetime.date.today().year,
    )


@app.route("/logout")
def logout():
    user_email = session.get("user", {}).get("email", "?")
    session.clear()
    log.info(f"Déconnexion : {user_email}")
    flash("Vous êtes déconnecté.", "success")
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Routes — Génération
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET", "POST"])
@login_required
def index():
    video_url = None
    video_info = None

    if request.method == "POST":
        file_gagne = request.files.get("file_gagne")
        file_pipe = request.files.get("file_pipe")
        ref_date_str = request.form.get("ref_date")

        if not file_gagne or not file_gagne.filename.endswith(".xlsx"):
            flash("Veuillez uploader un fichier Gagné .xlsx valide.", "error")
            return redirect(url_for("index"))
        if not file_pipe or not file_pipe.filename.endswith(".xlsx"):
            flash("Veuillez uploader un fichier Pipe .xlsx valide.", "error")
            return redirect(url_for("index"))

        # Sauvegarder les fichiers uploadés
        gagné_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}_gagne.xlsx")
        pipe_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}_pipe.xlsx")
        file_gagne.save(gagné_path)
        file_pipe.save(pipe_path)

        # Date de référence
        try:
            ref_date = datetime.date.fromisoformat(ref_date_str) if ref_date_str else datetime.date.today()
        except ValueError:
            ref_date = datetime.date.today()

        tmp_dir = None

        try:
            log.info(f"Génération lancée — date de référence : {ref_date}")

            # Pipeline de génération
            week_num = get_week_number(ref_date)
            monday, sunday = get_week_bounds(ref_date)
            week_dates = f"{monday.strftime('%d/%m')} — {sunday.strftime('%d/%m/%Y')}"

            # Événements depuis le fichier Gagné
            df_gagné = load_projects(gagné_path, label="Gagné")
            events_df = filter_events_this_week(df_gagné, ref_date)
            events = [format_event_card(row) for _, row in events_df.iterrows()]
            log.info(f"Semaine {week_num:02d} — {len(events)} événements trouvés")

            # Recos depuis le fichier Pipe
            df_pipe = load_projects(pipe_path, label="Pipe")
            recos_df = filter_recos_this_week(df_pipe, ref_date)
            recos = [format_reco_card(row) for _, row in recos_df.iterrows()]
            log.info(f"Semaine {week_num:02d} — {len(recos)} recos trouvées")

            if not events and not recos:
                flash(f"Aucun événement ni reco trouvé pour la semaine {week_num:02d} ({week_dates}).", "error")
                return redirect(url_for("index"))

            slides = build_slide_sequence(events, recos, week_num, week_dates)

            tmp_dir = tempfile.mkdtemp(prefix="sophie_slides_")
            image_durations = render_slides_to_images(slides, tmp_dir)

            video_filename = f"Semaine {week_num:02d} - Statistiques ecrans.mp4"
            output_path = os.path.join(OUTPUT_DIR, video_filename)
            build_video(image_durations, output_path)

            file_size = os.path.getsize(output_path) / 1024 / 1024

            # Upload vers Supabase Storage
            remote_path = upload_video(output_path, video_filename)
            if remote_path:
                video_url = get_video_url(remote_path)

            # Fallback : URL locale si Supabase indisponible
            if not video_url:
                video_url = url_for("download_video", filename=video_filename)

            video_info = f"Semaine {week_num:02d} — {len(events)} événements, {len(recos)} recos — {len(slides)} slides — {file_size:.1f} MB"

            log.info(f"Vidéo générée : {video_filename} ({file_size:.1f} MB)")
            flash("Vidéo générée avec succès !", "success")

        except ValueError as e:
            log.warning(f"Erreur de validation : {e}")
            flash(str(e), "error")
            return redirect(url_for("index"))

        except RuntimeError as e:
            log.error(f"Erreur de génération : {e}")
            flash(str(e), "error")
            return redirect(url_for("index"))

        except Exception as e:
            log.error(f"Erreur inattendue : {e}", exc_info=True)
            flash("Une erreur inattendue s'est produite. Voir les logs pour le détail.", "error")
            return redirect(url_for("index"))

        finally:
            # Nettoyage des fichiers uploadés
            for p in [gagné_path, pipe_path]:
                if os.path.exists(p):
                    os.remove(p)
            # Nettoyage du dossier temporaire de slides
            if tmp_dir and os.path.exists(tmp_dir):
                shutil.rmtree(tmp_dir, ignore_errors=True)

    user_email = session.get("user", {}).get("email") if _supabase_configured() else None

    return render_template_string(
        MAIN_TEMPLATE,
        today=datetime.date.today().isoformat(),
        year=datetime.date.today().year,
        video_url=video_url,
        video_info=video_info,
        user_email=user_email,
    )


@app.route("/download/<filename>")
@login_required
def download_video(filename):
    safe_name = secure_filename(filename)
    filepath = os.path.join(OUTPUT_DIR, safe_name)
    # Vérifier que le chemin résolu est bien dans OUTPUT_DIR
    if not os.path.realpath(filepath).startswith(os.path.realpath(OUTPUT_DIR)):
        flash("Accès non autorisé.", "error")
        return redirect(url_for("index"))
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=safe_name)
    flash("Fichier non trouvé.", "error")
    return redirect(url_for("index"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    log.info(f"Sophie Webapp démarrée sur http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
