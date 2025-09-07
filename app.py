from flask import Flask, request, send_file, render_template_string, redirect, url_for, flash
import msoffcrypto
import io

app = Flask(__name__)
app.secret_key = "une_cle_ultra_secrete"  # requis pour les flash messages

HTML_FORM = """
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Déverrouiller Excel</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    .toast-container { position: fixed; top: 1rem; right: 1rem; z-index: 2000; }
  </style>
</head>
<body class="bg-light">

  <div class="container py-5">
    <div class="row justify-content-center">
      <div class="col-md-6">
        <div class="card shadow-lg rounded-4">
          <div class="card-body p-4">
            <h2 class="text-center mb-4">Déverrouiller un fichier Excel</h2>
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
                <button type="submit" class="btn btn-primary btn-lg">Déverrouiller</button>
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

  <!-- Zone Toasts -->
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
    // Activer automatiquement tous les toasts
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
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=pwd)

            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)

            return send_file(
                decrypted,
                as_attachment=True,
                download_name=f.filename.replace(".xlsx", "_unlocked.xlsx")
            )
        except Exception:
            flash("Mot de passe incorrect ou fichier corrompu", "danger")
            return redirect(url_for("index"))

    return render_template_string(HTML_FORM)

if __name__ == "__main__":
    app.run(debug=True)
