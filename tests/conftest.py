# tests/conftest.py
import io
import re
import tempfile
from pathlib import Path

import pytest
from openpyxl import Workbook
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture()
def tmp_xlsm_path(tmp_path: Path):
    """Pad waar we een test-xlsm kunnen opslaan."""
    return tmp_path / "input.xlsm"


def _write_cell(ws, r, c, v):
    ws.cell(row=r, column=c, value=v)


@pytest.fixture()
def make_tabel_workbook():
    """
    Bouwt een minimale Excel die jouw parser verwacht:
    sheet 'Tabel' met anchor-rows + sample kolommen.
    """
    def _maker(
        *,
        project_code="T.25.16109",
        include_pfas=True,
        samples=("MM01", "02"),  # één mengmonster + één boorkop kolom
    ) -> bytes:
        wb = Workbook()
        ws = wb.active
        ws.title = "Tabel"

        # Zet ergens de projectcode
        _write_cell(ws, 1, 1, f"Project: {project_code}")

        # Anchor rows in kolom A (1)
        # Rij-indices zijn vrij, zolang anchors vindbaar zijn.
        row_mm = 5
        row_ms = 12
        row_rbk = 25
        row_pfas = 26
        row_tot = 27

        _write_cell(ws, row_mm, 1, "Mengmonster / boring")
        _write_cell(ws, row_ms, 1, "Monstersamenstelling")
        _write_cell(ws, row_rbk, 1, "Kwaliteitsklasse Rbk")
        _write_cell(ws, row_pfas, 1, "Kwaliteitsklasse PFAS")
        _write_cell(ws, row_tot, 1, "Kwaliteitsklasse totaal")

        # Sample kolommen starten vanaf kolom B (2)
        start_col = 2
        for idx, header in enumerate(samples):
            col = start_col + idx
            _write_cell(ws, row_mm, col, header)

            # Samenstelling (3 regels onder row_mm)
            _write_cell(ws, row_mm + 1, col, "Zand")
            _write_cell(ws, row_mm + 2, col, "Klei")
            _write_cell(ws, row_mm + 3, col, "Grind")

            # Boornummers blok onder Monstersamenstelling: vorm "01 (0,50-1,00)"
            # Parser loopt row_ms .. row_ms+12
            _write_cell(ws, row_ms + 1, col, "01 (0,50-1,00)")
            _write_cell(ws, row_ms + 2, col, "02 (0,50-1,00)")
            _write_cell(ws, row_ms + 3, col, "03 (0,50-1,00)")
            _write_cell(ws, row_ms + 4, col, "05 (0,50-1,00)")
            _write_cell(ws, row_ms + 5, col, "zz")  # moet genegeerd worden

            # Klassen (kolom-offset tolerant via _fetch_class)
            _write_cell(ws, row_rbk, col, "wonen")
            _write_cell(ws, row_pfas, col, "landbouw / natuur" if include_pfas else "")
            _write_cell(ws, row_tot, col, "klasse wonen")

        # PFAS rijen detectie: jouw parser zoekt keywords in col A
        if include_pfas:
            _write_cell(ws, 40, 1, "PFOS")  # bevat "PF"
            # zet een waarde in sample kolom zodat PFAS aanwezig is
            _write_cell(ws, 40, start_col, 0.12)

        # Arseen check: parser zoekt exact 'Arseen' in col A
        _write_cell(ws, 41, 1, "Arseen")
        _write_cell(ws, 41, start_col, 3.14)

        # Save naar bytes
        bio = io.BytesIO()
        wb.save(bio)
        return bio.getvalue()

    return _maker
