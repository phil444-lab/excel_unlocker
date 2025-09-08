from flask import Flask, request, send_file, render_template_string, redirect, url_for, flash
import msoffcrypto
import io
import os

app = Flask(__name__)
app.secret_key = "une_cle_ultra_secrete"

HTML_TEMPLATE = """
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Déverrouiller Excel</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body { padding-top: 70px; }
    footer { background: #f8f9fa; padding: 15px 0; margin-top: 40px; }
    .card { border-radius: 1rem; }
    .toast-container { position: fixed; top: 1rem; right: 1rem; z-index: 2000; }
  </style>
</head>
<body class="d-flex flex-column min-vh-100">

  <!-- Navbar -->
  <nav class="navbar navbar-expand-lg navbar-dark bg-success fixed-top shadow-sm">
    <div class="container">
      <a class="navbar-brand fw-bold" href="{{ url_for('index') }}">Excel Unlocker 🔓</a>
      <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarContent">
        <span class="navbar-toggler-icon"></span>
      </button>
      <div class="collapse navbar-collapse" id="navbarContent">
        <ul class="navbar-nav ms-auto">
          <li class="nav-item"><a class="nav-link active" href="{{ url_for('index') }}">Accueil</a></li>
          <li class="nav-item"><a class="nav-link" href="#">Aide</a></li>
        </ul>
      </div>
    </div>
  </nav>

  <!-- Contenu principal -->
  <main class="flex-fill">
    <div class="container py-5">
      <div class="row justify-content-center">
        <div class="col-lg-6 col-md-8">
          <div class="card shadow-lg">
            <div class="card-body p-4">
              <h2 class="text-center mb-4 text-success">Déverrouiller un fichier Excel</h2>
              <form method="post" enctype="multipart/form-data" class="needs-validation" novalidate>
                <div class="mb-3">
                  <label for="file" class="form-label">Fichier Excel protégé</label>
                  <input class="form-control" type="file" name="file" id="file" accept=".xlsx,.xls" required>
                </div>
                <div class="mb-3">
                  <label for="password" class="form-label">Mot de passe</label>
                  <input class="form-control" type="password" name="password" id="password" placeholder="Entrez le mot de passe" required>
                </div>
                <div class="d-grid">
                  <button type="submit" class="btn btn-success btn-lg">Déverrouiller & Télécharger</button>
                </div>
              </form>
            </div>
          </div>
          <p class="text-center text-muted small mt-3">
            ⚠️ Vos fichiers sont traités uniquement en mémoire, rien n’est stocké sur le serveur.
          </p>
        </div>
      </div>
    </div>
  </main>

  <!-- Footer -->
  <footer class="mt-auto">
    <div class="container text-center">
      <p class="mb-0 text-success">&copy; {{ year }} Excel Unlocker – Tous droits réservés</p>
    </div>
  </footer>

  <!-- Toasts -->
  <div class="toast-container">
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for category, msg in messages %}
          <div class="toast align-items-center text-bg-{{ category }} border-0 mb-2" role="alert" aria-live="assertive" aria-atomic="true">
            <div class="d-flex">
              <div class="toast-body">{{ msg }}</div>
              <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>
            </div>
          </div>
        {% endfor %}
      {% endif %}
    {% endwith %}
  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
  <script>
    document.querySelectorAll('.toast').forEach(toastEl => {
      new bootstrap.Toast(toastEl, { delay: 4000 }).show();
    });
  </script>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        f = request.files.get("file")
        pwd = request.form.get("password")

        if not f or not pwd:
            flash("Fichier ou mot de passe manquant", "warning")
            return redirect(url_for("index"))

        try:
            # Déchiffrement en mémoire
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=pwd)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)

            # Nom du fichier exporté
            xlsx_name = f.filename.rsplit(".", 1)[0] + "_unlocked.xlsx"

            return send_file(
                decrypted,
                as_attachment=True,
                download_name=xlsx_name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception:
            flash("Mot de passe incorrect ou fichier corrompu", "danger")
            return redirect(url_for("index"))

    from datetime import datetime
    return render_template_string(HTML_TEMPLATE, year=datetime.now().year)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
