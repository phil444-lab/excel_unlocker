from flask import Flask, request, send_file, render_template, redirect, url_for, flash, session
import msoffcrypto
import io
import os
import pandas as pd
from datetime import datetime
import tempfile

app = Flask(__name__)
app.secret_key = "une_cle_ultra_secrete"

# Dossier temporaire pour stocker les fichiers uploadés
UPLOAD_FOLDER = tempfile.gettempdir()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ---------------------------
# Déverrouillage
# ---------------------------
@app.route("/", methods=["GET", "POST"])
@app.route("/unlock", methods=["GET", "POST"])
def unlock():
    if request.method == "POST":
        f = request.files.get("file")
        pwd = request.form.get("password")
        if not f or not pwd:
            flash("Fichier ou mot de passe manquant", "warning")
            return redirect(url_for("unlock"))
        try:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=pwd)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            xlsx_name = f.filename.rsplit(".", 1)[0] + "_unlocked.xlsx"
            return send_file(
                decrypted,
                as_attachment=True,
                download_name=xlsx_name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception:
            flash("Mot de passe incorrect ou fichier corrompu", "danger")
            return redirect(url_for("unlock"))
    return render_template("unlock.html", year=datetime.now().year, active="unlock")

# ---------------------------
# Recherche avec possibilité de remplacer le fichier
# ---------------------------
@app.route("/search", methods=["GET", "POST"])
def search():
    results = {}
    query = ""
    file_name = session.get("file_name")
    sheet_names = session.get("sheet_names", [])
    selected_sheet = session.get("selected_sheet")
    file_path = session.get("file_path")

    if request.method == "POST":
        # Upload ou remplacement de fichier
        if "file" in request.files and request.files["file"].filename != "":
            f = request.files["file"]
            ext = f.filename.rsplit(".", 1)[-1].lower()
            if ext not in ["xlsx", "xls", "csv"]:
                flash("Format de fichier non supporté", "danger")
                return redirect(url_for("search"))

            # Sauvegarde temporaire côté serveur
            tmp_path = os.path.join(tempfile.gettempdir(), f.filename)
            f.save(tmp_path)

            # Remplacer l'ancien fichier si existant
            if file_path and os.path.exists(file_path) and file_path != tmp_path:
                try:
                    os.remove(file_path)
                except Exception:
                    pass

            # Mettre à jour la session
            session["file_path"] = tmp_path
            session["file_name"] = f.filename

            if ext in ["xlsx", "xls"]:
                xls = pd.ExcelFile(tmp_path)
                session["sheet_names"] = xls.sheet_names
                session["selected_sheet"] = xls.sheet_names[0]
            else:
                session["sheet_names"] = ["CSV"]
                session["selected_sheet"] = "CSV"

            flash(f"Fichier '{f.filename}' chargé avec succès.", "success")
            return redirect(url_for("search"))

        # Changement de feuille
        if "sheet" in request.form:
            selected_sheet = request.form["sheet"]
            session["selected_sheet"] = selected_sheet

        # Recherche
        if "query" in request.form:
            query = request.form["query"].strip()
            if not file_path or not query:
                flash("Aucun fichier chargé ou mot-clé vide", "warning")
                return redirect(url_for("search"))

            ext = file_path.rsplit(".", 1)[-1].lower()
            try:
                if ext in ["xlsx", "xls"]:
                    df = pd.read_excel(file_path, sheet_name=selected_sheet)
                else:
                    df = pd.read_csv(file_path)

                # Découpage des mots-clés
                keywords = query.split()

                # Recherche AND (toutes les occurrences dans la même ligne)
                mask_and = df.apply(
                    lambda row: all(
                        row.astype(str).str.contains(word, case=False, na=False).any()
                        for word in keywords
                    ),
                    axis=1
                )
                filtered = df[mask_and]

                # Si aucun résultat, on tente OR (au moins un mot)
                if filtered.empty:
                    mask_or = pd.Series(False, index=df.index)
                    for word in keywords:
                        mask_or |= df.apply(lambda row: row.astype(str).str.contains(word, case=False, na=False).any(), axis=1)
                    filtered = df[mask_or]

                if not filtered.empty:
                    filtered = filtered.fillna("Aucune")
                    
                    results[selected_sheet] = filtered.to_html(
                        classes="table table-sm table-bordered table-striped", index=False
                    )
                else:
                    flash("Aucun résultat trouvé", "info")
            except Exception as e:
                flash(f"Erreur lors de la lecture du fichier : {str(e)}", "danger")

    return render_template(
        "search.html",
        year=datetime.now().year,
        results=results,
        query=query,
        active="search",
        file_name=file_name,
        sheet_names=sheet_names,
        selected_sheet=selected_sheet
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
