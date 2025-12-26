# tests/test_app_generate.py
import re
from io import BytesIO

import pytest
from docx import Document

import app as terrapp  # verwacht: app.py exporteert `app` Flask instance


@pytest.fixture()
def client():
    terrapp.app.config["TESTING"] = True
    return terrapp.app.test_client()


def test_generate_requires_files(client):
    resp = client.post("/generate", data={}, content_type="multipart/form-data")
    assert resp.status_code in (400, 422)


def test_generate_returns_docx_with_project_filename(client, make_tabel_workbook, tmp_xlsm_path):
    xlsm_bytes = make_tabel_workbook(project_code="T.25.16109", include_pfas=True, samples=("MM01",))
    tmp_xlsm_path.write_bytes(xlsm_bytes)

    data = {
        "files": (open(tmp_xlsm_path, "rb"), "input.xlsm")
    }
    resp = client.post("/generate", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200
    assert resp.headers["Content-Type"].startswith(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    cd = resp.headers.get("Content-Disposition", "")
    # pas dit aan als jij 'T.25.16109_tabels.docx' of iets dergelijks gebruikt
    assert "T.25.16109" in cd
    assert cd.lower().endswith(".docx") or ".docx" in cd

    # check dat bytes een valide docx zijn
    doc = Document(BytesIO(resp.data))
    assert len(doc.tables) == 2
