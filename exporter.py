# exporter.py
from typing import List, Dict
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re

MC = 'Monstercode'
SAM = 'Samenstelling'
BN = 'Boornummer'
OND = 'Onderzochte parameters'
SKF = 'Stofspecifieke kwaliteitsklassen'
KKA = 'Kwaliteitsklasse analysemonster'

_CLASS_TOKEN_RE = re.compile(r'\b(L/N|W|IND|I|MV|SV)\b', re.IGNORECASE)


def _set_col_widths(table, widths_cm):
    """
    Zet kolombreedtes (cm) per kolom; schakelt autofit uit.
    widths_cm: list[float] met lengte == aantal kolommen
    """
    table.autofit = False
    ncols = len(table.columns)
    for j in range(min(ncols, len(widths_cm))):
        w = Cm(widths_cm[j])
        # zet breedte op alle cellen in deze kolom
        for row in table.rows:
            row.cells[j].width = w


def _cm_to_twips(cm: float) -> int:
    # 1 inch = 2.54 cm, 1 inch = 1440 twips
    return int(round(cm / 2.54 * 1440))


def _apply_fixed_layout(table):
    """
    Schakel AutoFit volledig uit door layout 'fixed' te zetten op de tabel.
    """
    table.autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr
    # verwijder bestaande tblLayout
    for el in tblPr.findall(qn('w:tblLayout')):
        tblPr.remove(el)
    # zet fixed layout
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'fixed')
    tblPr.append(tblLayout)


def _set_table_width(table, total_cm: float):
    """
    Zet de totale tabelbreedte (tblW) in twips.
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    for el in tblPr.findall(qn('w:tblW')):
        tblPr.remove(el)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:type'), 'dxa')
    tblW.set(qn('w:w'), str(_cm_to_twips(total_cm)))
    tblPr.append(tblW)


def _set_col_widths_strict(table, widths_cm):
    """
    Zet kolombreedtes dwingend:
    - tblGrid (gridCol)
    - elke cel tcW
    """
    # tblGrid
    tbl = table._tbl
    for el in tbl.findall(qn('w:tblGrid')):
        tbl.remove(el)
    grid = OxmlElement('w:tblGrid')
    widths_tw = [_cm_to_twips(w) for w in widths_cm]
    for tw in widths_tw:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(tw))
        grid.append(gridCol)
    tbl.append(grid)

    # elke cel dezelfde breedte per kolom
    for col_idx, tw in enumerate(widths_tw):
        for row in table.rows:
            cell = row.cells[col_idx]
            tcPr = cell._tc.get_or_add_tcPr()
            # verwijder bestaande tcW
            for el in tcPr.findall(qn('w:tcW')):
                tcPr.remove(el)
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:type'), 'dxa')
            tcW.set(qn('w:w'), str(tw))
            tcPr.append(tcW)


def _set_cell_font(cell, *, bold=False, italic=False, size_pt=9, color_rgb=None, name='Calibri'):
    for p in cell.paragraphs:
        for run in p.runs:
            run.font.name = name
            run.font.size = Pt(size_pt)
            run.bold = bold
            run.italic = italic
            if color_rgb:
                run.font.color.rgb = RGBColor.from_string(color_rgb)


def _add_table_heading(doc: Document, text: str):
    p = doc.add_paragraph(text)
    if not p.runs:
        p.add_run(text)
    for r in p.runs:
        r.font.name = 'Calibri'
        r.font.size = Pt(9)
        r.italic = True
    return p


def _add_note(doc: Document, text: str):
    p = doc.add_paragraph(text)
    if not p.runs:
        p.add_run(text)
    for r in p.runs:
        r.font.name = 'Calibri'
        r.font.size = Pt(8)
    return p


def _join_lines(items: List[str]) -> str:
    return '\n'.join(items)


def _set_cell_shading(cell, fill_hex: str):
    """Set cell background color (e.g., '008150')."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)


def _set_table_borders_horizontal_only(table):
    """
    Only horizontal black lines: top, bottom, insideH.
    Remove vertical borders (insideV) and left/right.
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    # Remove existing borders if any
    for el in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(el)

    borders = OxmlElement('w:tblBorders')

    def add_border(tag, val='single', sz='8', color='000000'):
        el = OxmlElement(tag)
        el.set(qn('w:val'), val)
        el.set(qn('w:sz'), sz)
        el.set(qn('w:color'), color)
        borders.append(el)

    # Top/Bottom strong, insideH thinner
    add_border('w:top',  'single', '8', '000000')
    add_border('w:bottom', 'single', '8', '000000')

    insideH = OxmlElement('w:insideH')
    insideH.set(qn('w:val'), 'single')
    insideH.set(qn('w:sz'), '4')
    insideH.set(qn('w:color'), '000000')
    borders.append(insideH)

    # Disable vertical and left/right
    for tag in ('w:insideV', 'w:left', 'w:right'):
        el = OxmlElement(tag)
        el.set(qn('w:val'), 'nil')
        borders.append(el)

    tblPr.append(borders)


def _extract_class_tokens(val: str) -> set:
    if not isinstance(val, str):
        return set()
    tokens = {t.upper() for t in _CLASS_TOKEN_RE.findall(val.upper())}
    if 'I' in tokens:
        tokens.discard('I')
        tokens.add('IND')  # normaliseer 'I' naar 'IND'
    return tokens


def export_to_docx(samples: List[Dict], out_path: str):
    # sorteer op monstercode (MM01..), fallback hoog
    def sort_key(s):
        try:
            return int(''.join(c for c in s[MC] if c.isdigit()))
        except:
            return 9999
    samples = sorted(samples, key=sort_key)

    doc = Document()

    # ===== Tabel 1 =====
    _add_table_heading(doc, "Tabel 1. Samenstelling analysemonsters.")
    cols1 = [MC, SAM, "Boornummer\n(traject in m - mv.)", "Onderzochte parameters"]
    t1 = doc.add_table(rows=1, cols=len(cols1))
    hdr1 = t1.rows[0].cells
    for j, name in enumerate(cols1):
        hdr1[j].text = name
        _set_cell_shading(hdr1[j], '008150')           # groene balk
        _set_cell_font(hdr1[j], bold=True, size_pt=9, color_rgb='FFFFFF')

    for s in samples:
        row = t1.add_row().cells
        row[0].text = s.get(MC, "")
        row[1].text = s.get(SAM, "")
        row[2].text = _join_lines(s.get(BN, []))
        row[3].text = s.get(OND, "")
        for c in row:
            _set_cell_font(c, size_pt=9)

    # [Monstercode, Samenstelling, Boornummer, Onderzochte parameters]
    widths_t1 = [2.13, 4.5, 3.75, 4.87]   # som ≈ 15.25 cm
    _apply_fixed_layout(t1)
    _set_table_width(t1, sum(widths_t1))
    _set_col_widths_strict(t1, widths_t1)
    _set_table_borders_horizontal_only(t1)

    _add_note(doc, "MM = mengmonster")
    _add_note(doc, "NEN 5740 grond:\t\tmetalen (barium, cadmium, kobalt, koper, kwik, lood, molybdeen, nikkel, zink), PAK (polycyclische \n\t\t\taromatische koolwaterstoffen), PCB (polychloorbifenylen), minerale olie, droge stof-, lutum- en \n\t\t\torganische stofgehalte.\nPFAS:\t\t\tper- en polyfluoralkylverbindingen")

    # ===== Tabel 2 =====
    _add_table_heading(doc, "Tabel 2. Samenvatting toetsing milieuhygiënische kwaliteit grond.")
    cols2 = [MC, SAM, "Boornummer\n(traject in m - mv.)", "Stofspecifieke kwaliteitsklassen", "Kwaliteitsklasse analysemonster"]
    t2 = doc.add_table(rows=1, cols=len(cols2))
    hdr2 = t2.rows[0].cells
    for j, name in enumerate(cols2):
        hdr2[j].text = name
        _set_cell_shading(hdr2[j], '008150')
        _set_cell_font(hdr2[j], bold=True, size_pt=9, color_rgb='FFFFFF')

    all_tokens = set()
    for s in samples:
        row = t2.add_row().cells
        row[0].text = s.get(MC, "")
        row[1].text = s.get(SAM, "")
        row[2].text = _join_lines(s.get(BN, []))
        skf_val = s.get(SKF, "")
        row[3].text = skf_val
        row[4].text = s.get(KKA, "")
        for c in row:
            _set_cell_font(c, size_pt=9)
        all_tokens |= _extract_class_tokens(skf_val)

    # [Monstercode, Samenstelling, Boornummer, SKF, KKA]
    widths_t2 = [2.13, 2.75, 3.75, 3.0, 3.5]   # som ≈ 15.13 cm
    _apply_fixed_layout(t2)
    _set_table_width(t2, sum(widths_t2))
    _set_col_widths_strict(t2, widths_t2)
    _set_table_borders_horizontal_only(t2)

    # Voorwaardelijke legenda
    only_ln = (len(all_tokens) > 0) and (all_tokens <= {'L/N'})
    if only_ln:
        _add_note(doc, "L/N : geen verontreinigingen aangetoond (de waarden overschrijden de kwaliteitseis voor klasse 'landbouw / natuur' niet)")
    else:
        _add_note(doc, "L/N\t: geen verontreinigingen aangetoond (de waarden overschrijden de kwaliteitseis voor klasse 'landbouw / natuur' niet)\nW\t: wonen (licht verontreinigd; de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'wonen')\nIND\t: industrie (licht verontreinigd; de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'industrie')\nMV\t: matig verontreinigd (de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'matig verontreinigd')\nSV\t: sterk verontreinigd (de aangetoonde waarden overschrijden de norm behorend bij de kwaliteitseis voor klasse 'matig verontreinigd' / interventiewaarde bodemkwaliteit (I))")

    doc.save(out_path)
