# tests/test_exporter_docx.py
import os
import tempfile
import zipfile
from io import BytesIO
from typing import Union

from docx import Document as DocxDocument

from exporter import export_to_docx


def _get_document_xml(docx_bytes: bytes) -> str:
    with zipfile.ZipFile(BytesIO(docx_bytes), "r") as z:
        return z.read("word/document.xml").decode("utf-8")


def _to_docx_bytes(ret) -> bytes:
    """
    Normalizeer export_to_docx return naar raw .docx bytes.
    Ondersteunt:
      - bytes
      - io.BytesIO
      - python-docx Document
      - str path naar bestand
    """
    # 1) bytes
    if isinstance(ret, (bytes, bytearray)):
        return bytes(ret)

    # 2) BytesIO / file-like met getvalue()
    if hasattr(ret, "getvalue"):
        try:
            b = ret.getvalue()
            if isinstance(b, (bytes, bytearray)):
                return bytes(b)
        except Exception:
            pass

    # 3) python-docx Document (heeft .save())
    if hasattr(ret, "save"):
        bio = BytesIO()
        ret.save(bio)
        bio.seek(0)
        return bio.read()

    # 4) pad als string
    if isinstance(ret, str) and os.path.exists(ret):
        with open(ret, "rb") as f:
            return f.read()

    raise TypeError(
        f"export_to_docx returned unsupported type: {type(ret)}. "
        "Make it return bytes/BytesIO/Document/or a file path."
    )


def test_exporter_generates_valid_docx_with_two_tables():
    samples = [
        {
            "Monstercode": "MM01",
            "Samenstelling": "Zand, Klei",
            "Boornummer": ["01 t/m 03 (0,50-1,00)", "05 (0,50-1,00)"],
            "Onderzochte parameters": "NEN 5740 grond, arseen, PFAS",
            "Stofspecifieke kwaliteitsklassen": "PFOS: L/N, Overig: W",
            "Kwaliteitsklasse analysemonster": "wonen",
        }
    ]

    ret = export_to_docx(samples)
    docx_bytes = _to_docx_bytes(ret)

    # 1) Word kan openen
    doc = DocxDocument(BytesIO(docx_bytes))
    assert len(doc.tables) == 2

    # 2) Check header teksten
    t1 = doc.tables[0]
    assert t1.rows[0].cells[0].text.strip().startswith("Monstercode")
    assert "Samenstelling" in t1.rows[0].cells[1].text

    # 3) XML checks: horizontale borders only
    xml = _get_document_xml(docx_bytes)
    assert "w:tblBorders" in xml
    assert "w:insideH" in xml
    assert 'w:insideV w:val="nil"' in xml or 'w:insideV w:val=\'nil\'' in xml
    assert 'w:left w:val="nil"' in xml or 'w:left w:val=\'nil\'' in xml
    assert 'w:right w:val="nil"' in xml or 'w:right w:val=\'nil\'' in xml

    # 4) Kolombreedtes aanwezig (tblGrid + tcW)
    assert "w:tblGrid" in xml
    assert "w:tcW" in xml

    # 5) Header shading groen
    assert 'w:fill="008150"' in xml or "w:fill='008150'" in xml

    # 6) Calibri mapping aanwezig
    assert "Calibri" in xml
