from flask import Flask, request, send_file, render_template_string
import msoffcrypto
import io

app = Flask(__name__)

HTML_FORM = """
<!doctype html>
<title>Déverrouiller Excel</title>
<h2>Upload un fichier Excel protégé</h2>
<form method=post enctype=multipart/form-data>
  <input type=file name=file required><br><br>
  Mot de passe: <input type=password name=password required><br><br>
  <input type=submit value="Déverrouiller">
</form>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        f = request.files["file"]
        pwd = request.form["password"]
        if not f or not pwd:
            return "Fichier ou mot de passe manquant", 400

        # Charger fichier en mémoire
        office_file = msoffcrypto.OfficeFile(f)
        try:
            office_file.load_key(password=pwd)
        except Exception:
            return "Mot de passe incorrect", 403

        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)
        decrypted.seek(0)

        return send_file(
            decrypted,
            as_attachment=True,
            download_name=f.filename.replace(".xlsx", "_unlocked.xlsx")
        )

    return render_template_string(HTML_FORM)

if __name__ == "__main__":
    app.run(debug=True)
