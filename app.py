# app.py
import io, os, tempfile
from flask import Flask, render_template, request, send_file, abort
from parser import parse_excel_to_samples
from exporter import export_to_docx
from docx import Document

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


# @app.route("/generate", methods=["POST"])
# def generate():
#     if "files" not in request.files:
#         abort(400, "Geen bestanden geüpload")
#     files = request.files.getlist("files")
#     all_samples = []
#     tmpdir = tempfile.mkdtemp()

#     for f in files:
#         if not f.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
#             continue
#         path = os.path.join(tmpdir, f.filename)
#         f.save(path)
#         try:
#             samples = parse_excel_to_samples(path)
#             all_samples.extend(samples)
#         except Exception as e:
#             print("Parse error:", e)

#     if not all_samples:
#         abort(400, "Geen bruikbare monsters gevonden in de uploads.")

#     out_io = io.BytesIO()
#     out_path = os.path.join(tmpdir, "tabels.docx")
#     export_to_docx(all_samples, out_path)
#     with open(out_path, "rb") as fh:
#         out_io.write(fh.read())
#     out_io.seek(0)
#     return send_file(
#         out_io,
#         as_attachment=True,
#         download_name="tabels.docx",
#         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#     )


@app.route("/generate", methods=["POST"])
def generate():
    if "files" not in request.files:
        abort(400, "Geen bestanden geüpload")
    files = request.files.getlist("files")
    all_samples = []
    project_codes = []
    tmpdir = tempfile.mkdtemp()

    for f in files:
        if not f.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            continue
        path = os.path.join(tmpdir, f.filename)
        f.save(path)
        try:
            samples, project_code = parse_excel_to_samples(path)
            all_samples.extend(samples)
            project_codes.append(project_code)
        except Exception as e:
            print("Parse error:", e)

    if not all_samples:
        abort(400, "Geen bruikbare monsters gevonden in de uploads.")

    out_io = export_to_docx(all_samples)
    project_code = project_codes[0]

    return send_file(
        out_io,
        as_attachment=True,
        download_name=f"{project_code}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        max_age=0,
        conditional=False,
    )


@app.get("/debug/docx-min")
def docx_min():
    bio = io.BytesIO()
    doc = Document()
    doc.add_paragraph("Hello from TerrApp – minimal docx")
    doc.save(bio)
    bio.seek(0)
    return send_file(
        bio,
        as_attachment=True,
        download_name="test.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        max_age=0,
        conditional=False,  # voorkom range/etag edge-cases
    )


if __name__ == "__main__":
    app.run(debug=True)
