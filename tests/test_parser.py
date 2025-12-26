# tests/test_parser.py
import re

from parser import parse_excel_to_samples


def test_parser_returns_samples_and_project_code(make_tabel_workbook, tmp_xlsm_path):
    xlsm_bytes = make_tabel_workbook(project_code="T.25.16109", include_pfas=True, samples=("MM01", "02"))
    tmp_xlsm_path.write_bytes(xlsm_bytes)

    samples, project_code = parse_excel_to_samples(str(tmp_xlsm_path))

    assert project_code == "T.25.16109"
    assert isinstance(samples, list)
    assert len(samples) == 2

    # Monstercodes
    codes = {s["Monstercode"] for s in samples}
    assert "MM01" in codes
    assert "02" in codes or "2" in codes  # afhankelijk van jouw normalisatie

    # Boor-groepering (01..03 samengevouwen + 05 apart)
    s0 = samples[0]
    boors = s0["Boornummer"]
    assert any("01 t/m 03" in line for line in boors)
    assert any("05" in line for line in boors)

    # Onderzochte parameters: NEN 5740 + arseen + PFAS
    assert "NEN 5740" in s0["Onderzochte parameters"]
    assert "arseen" in s0["Onderzochte parameters"]
    assert "PFAS" in s0["Onderzochte parameters"]
