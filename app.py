import io
import os
import shutil
import tempfile

from docx import Document
from flask import Flask, render_template, request, send_file

from parser import parse_excel_file
from exporter import export_to_docx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024


@app.route("/", methods=["GET"])
def index():
    last_update = get_last_update()
    return render_template("index.html", last_update=last_update)


@app.route("/generate", methods=["POST"])
def generate():
    if "files" not in request.files:
        last_update = get_last_update()
        return (
            render_template(
                "index.html",
                last_update=last_update,
                error_title="Bad Request",
                error_message="Geen bestanden geüpload.",
            ),
            400,
        )

    files = request.files.getlist("files")

    ground_samples = []
    groundwater_samples = []
    project_codes = []
    parse_errors = []

    tmpdir = tempfile.mkdtemp()

    try:
        for f in files:
            if not f.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
                continue

            safe_name = os.path.basename(f.filename)
            path = os.path.join(tmpdir, safe_name)
            f.save(path)

            try:
                g_samples, gw_samples, project_code = parse_excel_file(path)

                ground_samples.extend(g_samples)
                groundwater_samples.extend(gw_samples)

                if project_code:
                    project_codes.append(project_code)

            except Exception as e:
                msg = f"{safe_name}: {e}"
                parse_errors.append(msg)
                print("Parse error:", msg)

        if not ground_samples and not groundwater_samples:
            last_update = get_last_update()
            return (
                render_template(
                    "index.html",
                    last_update=last_update,
                    error_title="Bad Request",
                    error_message="Geen bruikbare grond- of grondwatertabellen gevonden in de uploads.",
                ),
                400,
            )

        out_io = export_to_docx(
            ground_samples=ground_samples,
            groundwater_samples=groundwater_samples,
        )

        project_code = project_codes[0] if project_codes else "export"

        return send_file(
            out_io,
            as_attachment=True,
            download_name=f"{project_code}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            max_age=0,
            conditional=False,
        )

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


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
        conditional=False,
    )


def get_last_update():
    ts = os.environ.get("BUILD_TIME")
    if ts:
        return ts
    return "onbekend"


if __name__ == "__main__":
    app.run(debug=True)