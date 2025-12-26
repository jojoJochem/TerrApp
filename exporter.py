import io
import re
from typing import List, Dict
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_ALIGN_VERTICAL

MC = 'Monstercode'
SAM = 'Samenstelling'
BN = 'Boornummer'
OND = 'Onderzochte parameters'
SKF = 'Stofspecifieke kwaliteitsklassen'
KKA = 'Kwaliteitsklasse analysemonster'

_CLASS_TOKEN_RE = re.compile(r'\b(L/N|W|IND|I|MV|SV)\b', re.IGNORECASE)

_L_N_LEGEND = "\nL/N\t: geen verontreinigingen aangetoond (de waarden overschrijden de kwaliteitseis voor klasse 'landbouw / natuur' niet)"
_W_LEGEND = "W\t: wonen (licht verontreinigd; de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'wonen')"
_IND_LEGEND = "IND\t: industrie (licht verontreinigd; de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'industrie')"
_MV_LEGEND = "MV\t: matig verontreinigd (de aangetoonde waarden voldoen aan de kwaliteitseis voor klasse 'matig verontreinigd')"
_SV_LEGEND = "SV\t: sterk verontreinigd (de aangetoonde waarden overschrijden de norm behorend bij de kwaliteitseis voor klasse 'matig verontreinigd' / interventiewaarde bodemkwaliteit (I))"

_MM_LEGEND = "\nMM = mengmonster"

_NEN_5740_LEGEND = (
        "NEN 5740 grond:\t\tmetalen (barium, cadmium, kobalt, koper, kwik, lood, molybdeen, nikkel, zink), PAK (polycyclische"
        "\n\t\t\taromatische koolwaterstoffen), PCB (polychloorbifenylen), minerale olie, droge stof-, lutum- en"
        "\n\t\t\torganische stofgehalte."
        "\nPFAS:\t\t\tper- en polyfluoralkylverbindingen"
    )

_LEGEND_FONT_SIZE = 8


def _dxa_from_cm(cm: float) -> int:
    # 1 inch = 2.54 cm, 1 inch = 1440 twips
    return int(round(cm / 2.54 * 1440))


def _force_calibri(run, size_pt=9, bold=False, italic=False, rgb=None):
    run.font.name = 'Calibri'
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    # hard-map rFonts zodat Windows niet terugvalt op Times
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), 'Calibri')
    rFonts.set(qn('w:hAnsi'), 'Calibri')
    rFonts.set(qn('w:cs'), 'Calibri')
    if rgb is not None:
        run.font.color.rgb = rgb


def _cell_runs_calibri(cell, size_pt=9):
    for p in cell.paragraphs:
        for r in p.runs:
            _force_calibri(r, size_pt=size_pt)


def _header_cell(cell, text, fill_hex="008150"):
    # achtergrond
    tcPr = cell._tc.get_or_add_tcPr()
    for el in tcPr.findall(qn('w:shd')):
        tcPr.remove(el)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)
    # tekst (multi-line), wit vet Calibri 9pt
    cell.text = ""
    for i, line in enumerate(str(text).split("\n")):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        _force_calibri(p.add_run(line), size_pt=9, bold=True, rgb=RGBColor(0xFF,0xFF,0xFF))
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER





def _borders_horizontal_only(table):
    """Alleen top/bottom/insideH, geen insideV/left/right (Windows-safe)."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    for el in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(el)
    borders = OxmlElement('w:tblBorders')

    # top/bottom stevig
    def add(tag, val='single', sz='8', color='000000'):
        el = OxmlElement(tag)
        el.set(qn('w:val'), val)
        el.set(qn('w:sz'), sz)
        el.set(qn('w:color'), color)
        borders.append(el)
    add('w:top')
    add('w:bottom')
    # insideH dunner
    insideH = OxmlElement('w:insideH')
    insideH.set(qn('w:val'), 'single')
    insideH.set(qn('w:sz'), '4')
    insideH.set(qn('w:color'), '000000')
    borders.append(insideH)
    # disable verticale & zijkanten
    for tag in ('w:insideV', 'w:left', 'w:right'):
        el = OxmlElement(tag)
        el.set(qn('w:val'), 'nil')
        borders.append(el)
    tblPr.append(borders)


def export_to_docx(samples: List[Dict]) -> io.BytesIO:
    # sortering op monstercode
    def sort_key(s):
        try:
            return int(''.join(c for c in s.get(MC,'') if c.isdigit()))
        except:
            return 9999
    samples = sorted(samples, key=sort_key)

    doc = Document()
    # style "Normal" richting Calibri 9pt
    try:
        normal = doc.styles['Normal']
        normal.font.name = 'Calibri'
        normal.font.size = Pt(9)
        n = normal._element
        rPr = n.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
        if rFonts.getparent() is None:
            rPr.append(rFonts)
        for k in ('w:ascii', 'w:hAnsi', 'w:cs'): rFonts.set(qn(k), 'Calibri')
    except Exception:
        pass

    # ===== Tabel 1 =====
    p1 = doc.add_paragraph("Tabel 1. Samenstelling analysemonsters.")
    for r in p1.runs:
        _force_calibri(r, size_pt=9, italic=True)

    cols1 = [MC, SAM, "Boornummer\n(traject in m - mv.)", OND]
    t1 = doc.add_table(rows=1, cols=len(cols1))
    t1.style = "Table Grid"
    for j, name in enumerate(cols1):
        _header_cell(t1.rows[0].cells[j], name)

    for s in samples:
        row = t1.add_row().cells
        row[0].text = s.get(MC, "")
        row[1].text = s.get(SAM, "")
        row[2].text = ""
        lines = s.get(BN, []) or []
        if not isinstance(lines, list):
            lines = [str(lines)]
        for i, line in enumerate(lines):
            (_p := (row[2].paragraphs[0] if i == 0 else row[2].add_paragraph())).add_run(str(line))
        row[3].text = s.get(OND, "")
        for c in row:
            _cell_runs_calibri(c, size_pt=9)

    _borders_horizontal_only(t1)

    # notities
    n1 = doc.add_paragraph(_MM_LEGEND)
    [_force_calibri(r, size_pt=_LEGEND_FONT_SIZE) for r in n1.runs]
    n2 = doc.add_paragraph(_NEN_5740_LEGEND)
    [_force_calibri(r, size_pt=_LEGEND_FONT_SIZE) for r in n2.runs]

    # ===== Tabel 2 =====
    p2 = doc.add_paragraph("Tabel 2. Samenvatting toetsing milieuhygiënische kwaliteit grond.")
    for r in p2.runs:
        _force_calibri(r, size_pt=9, italic=True)

    cols2 = [MC, SAM, "Boornummer\n(traject in m - mv.)", SKF, KKA]
    t2 = doc.add_table(rows=1, cols=len(cols2))
    t2.style = "Table Grid"
    for j, name in enumerate(cols2):
        _header_cell(t2.rows[0].cells[j], name)

    tokens = set()
    for s in samples:
        row = t2.add_row().cells
        row[0].text = s.get(MC, "")
        row[1].text = s.get(SAM, "")
        row[2].text = ""
        lines = s.get(BN, []) or []
        if not isinstance(lines, list):
            lines = [str(lines)]
        for i, line in enumerate(lines):
            (_p := (row[2].paragraphs[0] if i == 0 else row[2].add_paragraph())).add_run(str(line))
        skf_val = (s.get(SKF, "") or "")
        row[3].text = skf_val
        row[4].text = s.get(KKA, "") or ""
        for c in row:
            _cell_runs_calibri(c, size_pt=9)
        for tok in _CLASS_TOKEN_RE.findall(skf_val.upper()):
            tokens.add("IND" if tok == "I" else tok)

    _borders_horizontal_only(t2)

    # legenda
    only_ln = bool(tokens) and tokens <= {"L/N"}
    leg = _L_N_LEGEND if only_ln else f"{_L_N_LEGEND}\n{_W_LEGEND}\n{_IND_LEGEND}\n{_MV_LEGEND}\n{_SV_LEGEND}"
    p_leg = doc.add_paragraph(leg)
    [_force_calibri(r, size_pt=_LEGEND_FONT_SIZE) for r in p_leg.runs]

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
