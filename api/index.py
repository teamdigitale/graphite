from flask import Flask, request, Response, render_template, send_file
import os
import pandas as pd
from docx import Document
from werkzeug.utils import secure_filename
import zipfile
from functools import wraps
import base64
import logging

logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
app.secret_key = "secret"
UPLOAD_FOLDER = "/tmp"

# --- AUTENTICAZIONE BASIC ---
def check_auth(auth_header: str) -> bool:
    """Verifica header Basic Auth rispetto a variabili ambiente BASIC_AUTH_PASSWORDS"""
    if not auth_header or not auth_header.startswith("Basic "):
        return False

    try:
        encoded = auth_header.split(" ", 1)[1].strip()
        decoded = base64.b64decode(encoded).decode("utf-8")
        user_input, pwd_input = decoded.split(":", 1)
    except Exception:
        return False

    valid_pairs = os.environ.get("BASIC_AUTH_PASSWORDS", "")
    for pair in valid_pairs.strip().split():
        if ":" in pair:
            valid_user, valid_pwd = pair.split(":", 1)
            if user_input == valid_user and pwd_input == valid_pwd:
                return True
    return False

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth_header = request.headers.get("Authorization")
        if not check_auth(auth_header):
            return Response(
                "Autenticazione richiesta", 401,
                {"WWW-Authenticate": 'Basic realm="Login Required"'}
            )
        return f(*args, **kwargs)
    return decorated

# --- GENERAZIONE DOCUMENTI ---
def generate_documents(excel_path, word_path, prefix, selected_rows):
    logging.debug("Inizio generazione documenti")
    
    try:
        df = pd.read_excel(excel_path)
        logging.debug(f"File Excel caricato: {excel_path}")
    except Exception as e:
        logging.error(f"Errore nel caricamento del file Excel: {e}")
        raise

    output_dir = os.path.join(UPLOAD_FOLDER, "output_docs")
    os.makedirs(output_dir, exist_ok=True)
    output_files = []

    for idx in selected_rows:
        if idx >= len(df):
            logging.warning(f"Riga {idx} fuori intervallo")
            continue
        row = df.iloc[idx]
        try:
            doc = Document(word_path)
            logging.debug(f"Documento Word caricato: {word_path}")
        except Exception as e:
            logging.error(f"Errore nel caricamento del documento Word: {e}")
            raise

        # Sostituzione nei paragrafi mantenendo il formato
        try:
            for paragraph in doc.paragraphs:
                for key, value in row.items():
                    if f"{{{{{key}}}}}" in paragraph.text:
                        for run in paragraph.runs:
                            if f"{{{{{key}}}}}" in run.text:
                                original_font = run.font
                                run.text = run.text.replace(f"{{{{{key}}}}}", str(value))
                                # Ripristina lo stile originale del font
                                run.font.name = original_font.name
                                run.font.size = original_font.size
                                run.font.bold = original_font.bold
                                run.font.italic = original_font.italic
                                run.font.underline = original_font.underline
                                run.font.color.rgb = original_font.color.rgb
        except Exception as e:
            logging.error(f"Errore durante la sostituzione nel paragrafo: {e}")
            raise

        # Sostituzione nelle tabelle mantenendo il formato
        try:
            for table in doc.tables:
                for row_table in table.rows:
                    for cell in row_table.cells:
                        for key, value in row.items():
                            if f"{{{{{key}}}}}" in cell.text:
                                for paragraph in cell.paragraphs:
                                    for run in paragraph.runs:
                                        if f"{{{{{key}}}}}" in run.text:
                                            original_font = run.font
                                            run.text = run.text.replace(f"{{{{{key}}}}}", str(value))
                                            # Ripristina lo stile originale del font
                                            run.font.name = original_font.name
                                            run.font.size = original_font.size
                                            run.font.bold = original_font.bold
                                            run.font.italic = original_font.italic
                                            run.font.underline = original_font.underline
                                            run.font.color.rgb = original_font.color.rgb
        except Exception as e:
            logging.error(f"Errore durante la sostituzione nella tabella: {e}")
            raise

        filename = f"{prefix}{row.iloc[0]}_{idx}.docx"
        filepath = os.path.join(output_dir, filename)
        try:
            doc.save(filepath)
            output_files.append(filepath)
        except Exception as e:
            logging.error(f"Errore durante il salvataggio del documento: {e}")
            raise

    return output_files

def parse_row_selection(range_rows, specific_rows, total_rows):
    selected = set()

    # Intervallo tipo "2-10"
    if range_rows:
        try:
            start, end = map(int, range_rows.split("-"))
            selected.update(range(start - 1, end))
        except Exception:
            pass

    # Righe specifiche tipo "3,7,9"
    if specific_rows:
        try:
            selected.update(int(i) - 1 for i in specific_rows.split(",") if i.strip().isdigit())
        except Exception:
            pass

    return selected if selected else range(total_rows)

# --- ROUTE PRINCIPALE ---
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
            return "File Excel o Word mancante.", 400

        excel_path = os.path.join(UPLOAD_FOLDER, secure_filename(excel.filename))
        word_path = os.path.join(UPLOAD_FOLDER, secure_filename(word.filename))
        excel.save(excel_path)
        word.save(word_path)

        ext = os.path.splitext(excel_path)[1].lower()
        if ext == ".xls":
            df = pd.read_excel(excel_path, engine="xlrd")
        else:
            df = pd.read_excel(excel_path, engine="openpyxl")

        selected_rows = parse_row_selection(range_rows, specific_rows, len(df))
        output_files = generate_documents(excel_path, word_path, prefix, selected_rows)

        zip_path = os.path.join(UPLOAD_FOLDER, "output.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file in output_files:
                zipf.write(file, os.path.basename(file))

        return send_file(zip_path, as_attachment=True)

    return render_template("upload.html")