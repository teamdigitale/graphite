from flask import Flask, request, Response, render_template, send_file
import os
import pandas as pd
from docx import Document
from werkzeug.utils import secure_filename
import zipfile
from functools import wraps

app = Flask(__name__)
app.secret_key = "secret"
UPLOAD_FOLDER = "/tmp"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def check_auth(username, password):
    return (
        username == os.environ.get("BASIC_AUTH_USERNAME")
        and password == os.environ.get("BASIC_AUTH_PASSWORD")
    )

def authenticate():
    return Response(
        "Accesso richiesto", 401,
        {"WWW-Authenticate": 'Basic realm="Login Required"'}
    )

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

@app.route("/", methods=["GET", "POST"])
@requires_auth
def upload():
    if request.method == "POST":
        excel = request.files.get("excel")
        word = request.files.get("word")
        prefix = request.form.get("prefix", "")
        range_rows = request.form.get("range_rows", "")
        specific_rows = request.form.get("specific_rows", "")

        if not excel or not word:
            return "File mancante", 400

        excel_path = os.path.join(UPLOAD_FOLDER, secure_filename(excel.filename))
        word_path = os.path.join(UPLOAD_FOLDER, secure_filename(word.filename))
        excel.save(excel_path)
        word.save(word_path)

        df = pd.read_excel(excel_path)
        rows_to_process = set()

        if range_rows:
            try:
                start, end = map(int, range_rows.split("-"))
                rows_to_process.update(range(start - 1, end))
            except:
                pass

        if specific_rows:
            try:
                rows_to_process.update(int(i)-1 for i in specific_rows.split(","))
            except:
                pass

        if not rows_to_process:
            rows_to_process = range(len(df))

        output_dir = os.path.join(UPLOAD_FOLDER, "output_docs")
        os.makedirs(output_dir, exist_ok=True)
        output_files = []

        for idx in rows_to_process:
            if idx >= len(df):
                continue
            row = df.iloc[idx]
            doc = Document(word_path)
            for paragraph in doc.paragraphs:
                for key, value in row.items():
                    paragraph.text = paragraph.text.replace(f"{{{{{key}}}}}", str(value))
            for table in doc.tables:
                for row_table in table.rows:
                    for cell in row_table.cells:
                        for key, value in row.items():
                            cell.text = cell.text.replace(f"{{{{{key}}}}}", str(value))
            filename = f"{prefix}{row.iloc[0]}_{idx}.docx"
            filepath = os.path.join(output_dir, filename)
            doc.save(filepath)
            output_files.append(filepath)

        zip_path = os.path.join(UPLOAD_FOLDER, "output.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file in output_files:
                zipf.write(file, os.path.basename(file))

        return send_file(zip_path, as_attachment=True)

    return render_template("upload.html")
