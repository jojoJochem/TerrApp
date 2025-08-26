# # exporter.py
# from typing import List, Dict
# from docx import Document
# from docx.shared import Pt, RGBColor, Cm
# from docx.oxml import OxmlElement
# from docx.oxml.ns import qn
# import re
# import io
# from typing import List, Dict
# from docx import Document

# MC = 'Monstercode'
# SAM = 'Samenstelling'
# BN = 'Boornummer'
# OND = 'Onderzochte parameters'
# SKF = 'Stofspecifieke kwaliteitsklassen'
# KKA = 'Kwaliteitsklasse analysemonster'

# _CLASS_TOKEN_RE = re.compile(r'\b(L/N|W|IND|I|MV|SV)\b', re.IGNORECASE)


# def _set_col_widths(table, widths_cm):
#     """
#     Zet kolombreedtes (cm) per kolom; schakelt autofit uit.
#     widths_cm: list[float] met lengte == aantal kolommen
#     """
#     table.autofit = False
#     ncols = len(table.columns)
#     for j in range(min(ncols, len(widths_cm))):
#         w = Cm(widths_cm[j])
#         # zet breedte op alle cellen in deze kolom
#         for row in table.rows:
#             row.cells[j].width = w


# def _cm_to_twips(cm: float) -> int:
#     # 1 inch = 2.54 cm, 1 inch = 1440 twips
#     return int(round(cm / 2.54 * 1440))


# def _apply_fixed_layout(table):
#     """
#     Schakel AutoFit volledig uit door layout 'fixed' te zetten op de tabel.
#     """
#     table.autofit = False
#     tbl = table._tbl
#     tblPr = tbl.tblPr
#     # verwijder bestaande tblLayout
#     for el in tblPr.findall(qn('w:tblLayout')):
#         tblPr.remove(el)
#     # zet fixed layout
#     tblLayout = OxmlElement('w:tblLayout')
#     tblLayout.set(qn('w:type'), 'fixed')
#     tblPr.append(tblLayout)


# def _set_table_width(table, total_cm: float):
#     """
#     Zet de totale tabelbreedte (tblW) in twips.
#     """
#     tbl = table._tbl
#     tblPr = tbl.tblPr
#     for el in tblPr.findall(qn('w:tblW')):
#         tblPr.remove(el)
#     tblW = OxmlElement('w:tblW')
#     tblW.set(qn('w:type'), 'dxa')
#     tblW.set(qn('w:w'), str(_cm_to_twips(total_cm)))
#     tblPr.append(tblW)


# def _set_col_widths_strict(table, widths_cm):
#     """
#     Zet kolombreedtes dwingend:
#     - tblGrid (gridCol)
#     - elke cel tcW
#     """
#     # tblGrid
#     tbl = table._tbl
#     for el in tbl.findall(qn('w:tblGrid')):
#         tbl.remove(el)
#     grid = OxmlElement('w:tblGrid')
#     widths_tw = [_cm_to_twips(w) for w in widths_cm]
#     for tw in widths_tw:
#         gridCol = OxmlElement('w:gridCol')
#         gridCol.set(qn('w:w'), str(tw))
#         grid.append(gridCol)
#     tbl.append(grid)

#     # elke cel dezelfde breedte per kolom
#     for col_idx, tw in enumerate(widths_tw):
#         for row in table.rows:
#             cell = row.cells[col_idx]
#             tcPr = cell._tc.get_or_add_tcPr()
#             # verwijder bestaande tcW
#             for el in tcPr.findall(qn('w:tcW')):
#                 tcPr.remove(el)
#             tcW = OxmlElement('w:tcW')
#             tcW.set(qn('w:type'), 'dxa')
#             tcW.set(qn('w:w'), str(tw))
#             tcPr.append(tcW)


# def _set_cell_font(cell, *, bold=False, italic=False, size_pt=9, color_rgb=None, name='Calibri'):
#     for p in cell.paragraphs:
#         for run in p.runs:
#             run.font.name = name
#             run.font.size = Pt(size_pt)
#             run.bold = bold
#             run.italic = italic
#             if color_rgb:
#                 run.font.color.rgb = RGBColor.from_string(color_rgb)


# def _add_table_heading(doc: Document, text: str):
#     p = doc.add_paragraph(text)
#     if not p.runs:
#         p.add_run(text)
#     for r in p.runs:
#         r.font.name = 'Calibri'
#         r.font.size = Pt(9)
#         r.italic = True
#     return p


# def _add_note(doc: Document, text: str):
#     p = doc.add_paragraph(text)
#     if not p.runs:
#         p.add_run(text)
#     for r in p.runs:
#         r.font.name = 'Calibri'
#         r.font.size = Pt(8)
#     return p


# def _join_lines(items: List[str]) -> str:
#     return '\n'.join(items)


# def _set_cell_shading(cell, fill_hex: str):
#     """Set cell background color (e.g., '008150')."""
#     tc = cell._tc
#     tcPr = tc.get_or_add_tcPr()
#     shd = OxmlElement('w:shd')
#     shd.set(qn('w:val'), 'clear')
#     shd.set(qn('w:color'), 'auto')
#     shd.set(qn('w:fill'), fill_hex)
#     tcPr.append(shd)


# def _set_table_borders_horizontal_only(table):
#     """
#     Only horizontal black lines: top, bottom, insideH.
#     Remove vertical borders (insideV) and left/right.
#     """
#     tbl = table._tbl
#     tblPr = tbl.tblPr
#     # Remove existing borders if any
#     for el in tblPr.findall(qn('w:tblBorders')):
#         tblPr.remove(el)

#     borders = OxmlElement('w:tblBorders')

#     def add_border(tag, val='single', sz='8', color='000000'):
#         el = OxmlElement(tag)
#         el.set(qn('w:val'), val)
#         el.set(qn('w:sz'), sz)
#         el.set(qn('w:color'), color)
#         borders.append(el)

#     # Top/Bottom strong, insideH thinner
#     add_border('w:top',  'single', '8', '000000')
#     add_border('w:bottom', 'single', '8', '000000')

#     insideH = OxmlElement('w:insideH')
#     insideH.set(qn('w:val'), 'single')
#     insideH.set(qn('w:sz'), '4')
#     insideH.set(qn('w:color'), '000000')
#     borders.append(insideH)

#     # Disable vertical and left/right
#     for tag in ('w:insideV', 'w:left', 'w:right'):
#         el = OxmlElement(tag)
#         el.set(qn('w:val'), 'nil')
#         borders.append(el)

#     tblPr.append(borders)


# def _extract_class_tokens(val: str) -> set:
#     if not isinstance(val, str):
#         return set()
#     tokens = {t.upper() for t in _CLASS_TOKEN_RE.findall(val.upper())}
#     if 'I' in tokens:
#         tokens.discard('I')
#         tokens.add('IND')  # normaliseer 'I' naar 'IND'
#     return tokens


# # def export_to_docx(samples: List[Dict], out_path: str):
# #     # sorteer op monstercode (MM01..), fallback hoog
# #     def sort_key(s):
# #         try:
# #             return int(''.join(c for c in s[MC] if c.isdigit()))
# #         except:
# #             return 9999
# #     samples = sorted(samples, key=sort_key)

# #     doc = Document()

# #     # ===== Tabel 1 =====
# #     _add_table_heading(doc, "Tabel 1. Samenstelling analysemonsters.")
# #     cols1 = [MC, SAM, "Boornummer\n(traject in m - mv.)", "Onderzochte parameters"]
# #     t1 = doc.add_table(rows=1, cols=len(cols1))
# #     hdr1 = t1.rows[0].cells
# #     for j, name in enumerate(cols1):
# #         hdr1[j].text = name
# #         _set_cell_shading(hdr1[j], '008150')           # groene balk
# #         _set_cell_font(hdr1[j], bold=True, size_pt=9, color_rgb='FFFFFF')

# #     for s in samples:
# #         row = t1.add_row().cells
# #         row[0].text = s.get(MC, "")
# #         row[1].text = s.get(SAM, "")
# #         row[2].text = _join_lines(s.get(BN, []))
# #         row[3].text = s.get(OND, "")
# #         for c in row:
# #             _set_cell_font(c, size_pt=9)

# #     # [Monstercode, Samenstelling, Boornummer, Onderzochte parameters]
# #     widths_t1 = [2.13, 4.5, 3.75, 4.87]   # som ≈ 15.25 cm
# #     _apply_fixed_layout(t1)
# #     _set_table_width(t1, sum(widths_t1))
# #     _set_col_widths_strict(t1, widths_t1)
# #     _set_table_borders_horizontal_only(t1)

# #     _add_note(doc, "MM = mengmonster")
# #     _add_note(doc, "NEN 5740 grond:\t\tmetalen (barium, cadmium, kobalt, koper, kwik, lood, molybdeen, nikkel, zink), PAK (polycyclische \n\t\t\taromatische koolwaterstoffen), PCB (polychloorbifenylen), minerale olie, droge stof-, lutum- en \n\t\t\torganische stofgehalte.\nPFAS:\t\t\tper- en polyfluoralkylverbindingen")

# #     # ===== Tabel 2 =====
# #     _add_table_heading(doc, "Tabel 2. Samenvatting toetsing milieuhygiënische kwaliteit grond.")
# #     cols2 = [MC, SAM, "Boornummer\n(traject in m - mv.)", "Stofspecifieke kwaliteitsklassen", "Kwaliteitsklasse analysemonster"]
# #     t2 = doc.add_table(rows=1, cols=len(cols2))
# #     hdr2 = t2.rows[0].cells
# #     for j, name in enumerate(cols2):
# #         hdr2[j].text = name
# #         _set_cell_shading(hdr2[j], '008150')
# #         _set_cell_font(hdr2[j], bold=True, size_pt=9, color_rgb='FFFFFF')

# #     all_tokens = set()
# #     for s in samples:
# #         row = t2.add_row().cells
# #         row[0].text = s.get(MC, "")
# #         row[1].text = s.get(SAM, "")
# #         row[2].text = _join_lines(s.get(BN, []))
# #         skf_val = s.get(SKF, "")
# #         row[3].text = skf_val
# #         row[4].text = s.get(KKA, "")
# #         for c in row:
# #             _set_cell_font(c, size_pt=9)
# #         all_tokens |= _extract_class_tokens(skf_val)

# #     # [Monstercode, Samenstelling, Boornummer, SKF, KKA]
# #     widths_t2 = [2.13, 2.75, 3.75, 3.0, 3.5]   # som ≈ 15.13 cm
# #     _apply_fixed_layout(t2)
# #     _set_table_width(t2, sum(widths_t2))
# #     _set_col_widths_strict(t2, widths_t2)
# #     _set_table_borders_horizontal_only(t2)

# #     # Voorwaardelijke legenda
# #     only_ln = (len(all_tokens) > 0) and (all_tokens <= {'L/N'})
# #     if only_ln:
# #         _add_note(doc, "L/N : geen verontreinigingen aangetoond (de waarden overschrijden de kwaliteitseis voor klasse 'landbouw / natuur' niet)")
# #     else:
# #         _add_note(doc, "L/N\t: geen verontreinigingen aangetoond (de waarden overschrijden de kwaliteitseis voor klasse 'landbouw / natuur' niet)\nW\t: wonen (licht verontreinigd; de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'wonen')\nIND\t: industrie (licht verontreinigd; de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'industrie')\nMV\t: matig verontreinigd (de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'matig verontreinigd')\nSV\t: sterk verontreinigd (de aangetoonde waarden overschrijden de norm behorend bij de kwaliteitseis voor klasse 'matig verontreinigd' / interventiewaarde bodemkwaliteit (I))")

# #     doc.save(out_path)


# def export_to_docx(samples: List[Dict], out_path: str | None = None):
#     """
#     Bouwt het Word-document.
#     - Als out_path is opgegeven: schrijft naar bestand en returnt het pad.
#     - Als out_path None is: returnt een BytesIO met de .docx-inhoud (seek(0) gezet).
#     """
#     # sorteer op monstercode (MM01..), fallback hoog
#     def sort_key(s):
#         try:
#             return int(''.join(c for c in s['Monstercode'] if c.isdigit()))
#         except:
#             return 9999
#     samples = sorted(samples, key=sort_key)

#     doc = Document()

#     # ===== Tabel 1 =====
#     _add_table_heading(doc, "Tabel 1. Samenstelling analysemonsters.")
#     cols1 = ['Monstercode', 'Samenstelling', "Boornummer\n(traject in m - mv.)", "Onderzochte parameters"]
#     t1 = doc.add_table(rows=1, cols=len(cols1))
#     hdr1 = t1.rows[0].cells
#     for j, name in enumerate(cols1):
#         hdr1[j].text = name
#         _set_cell_shading(hdr1[j], '008150')
#         _set_cell_font(hdr1[j], bold=True, size_pt=9, color_rgb='FFFFFF')

#     for s in samples:
#         row = t1.add_row().cells
#         row[0].text = s.get('Monstercode', "")
#         row[1].text = s.get('Samenstelling', "")
#         row[2].text = '\n'.join(s.get('Boornummer', []))
#         row[3].text = s.get('Onderzochte parameters', "")
#         for c in row:
#             _set_cell_font(c, size_pt=9)

#     widths_t1 = [2.13, 4.5, 3.75, 4.87]
#     _apply_fixed_layout(t1)
#     _set_table_width(t1, sum(widths_t1))
#     _set_col_widths_strict(t1, widths_t1)
#     _set_table_borders_horizontal_only(t1)

#     _add_note(doc, "MM = mengmonster")
#     _add_note(doc, "NEN 5740 grond:\t\tmetalen (barium, cadmium, kobalt, koper, kwik, lood, molybdeen, nikkel, zink), PAK (polycyclische \n\t\t\taromatische koolwaterstoffen), PCB (polychloorbifenylen), minerale olie, droge stof-, lutum- en \n\t\t\torganische stofgehalte.\nPFAS:\t\t\tper- en polyfluoralkylverbindingen")

#     # ===== Tabel 2 =====
#     _add_table_heading(doc, "Tabel 2. Samenvatting toetsing milieuhygiënische kwaliteit grond.")
#     cols2 = ['Monstercode', 'Samenstelling', "Boornummer\n(traject in m - mv.)", "Stofspecifieke kwaliteitsklassen", "Kwaliteitsklasse analysemonster"]
#     t2 = doc.add_table(rows=1, cols=len(cols2))
#     hdr2 = t2.rows[0].cells
#     for j, name in enumerate(cols2):
#         hdr2[j].text = name
#         _set_cell_shading(hdr2[j], '008150')
#         _set_cell_font(hdr2[j], bold=True, size_pt=9, color_rgb='FFFFFF')

#     all_tokens = set()
#     for s in samples:
#         row = t2.add_row().cells
#         row[0].text = s.get('Monstercode', "")
#         row[1].text = s.get('Samenstelling', "")
#         row[2].text = '\n'.join(s.get('Boornummer', []))
#         skf_val = s.get('Stofspecifieke kwaliteitsklassen', "")
#         row[3].text = skf_val
#         row[4].text = s.get('Kwaliteitsklasse analysemonster', "")
#         for c in row:
#             _set_cell_font(c, size_pt=9)
#         all_tokens |= _extract_class_tokens(skf_val)

#     widths_t2 = [2.13, 2.75, 3.75, 3.0, 3.5]
#     _apply_fixed_layout(t2)
#     _set_table_width(t2, sum(widths_t2))
#     _set_col_widths_strict(t2, widths_t2)
#     _set_table_borders_horizontal_only(t2)

#     only_ln = (len(all_tokens) > 0) and (all_tokens <= {'L/N'})
#     if only_ln:
#         _add_note(doc, "L/N : geen verontreinigingen aangetoond (de waarden overschrijden de kwaliteitseis voor klasse 'landbouw / natuur' niet)")
#     else:
#         _add_note(doc, "L/N\t: geen verontreinigingen aangetoond (de waarden overschrijden de kwaliteitseis voor klasse 'landbouw / natuur' niet)\nW\t: wonen (licht verontreinigd; de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'wonen')\nIND\t: industrie (licht verontreinigd; de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'industrie')\nMV\t: matig verontreinigd (de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'matig verontreinigd')\nSV\t: sterk verontreinigd (de aangetoonde waarden overschrijden de norm behorend bij de kwaliteitseis voor klasse 'matig verontreinigd' / interventiewaarde bodemkwaliteit (I))")

#     if out_path:
#         doc.save(out_path)
#         return out_path
#     else:
#         bio = io.BytesIO()
#         doc.save(bio)
#         bio.seek(0)
#         return bio



# def export_to_docx_win_safe(samples: List[Dict]) -> io.BytesIO:
#     """Windows-compatibele .docx zonder low-level XML table tweaks."""
#     # sorteer op monstercode (MM01..)
#     def sort_key(s):
#         try:    return int(''.join(c for c in s.get(MC,'') if c.isdigit()))
#         except: return 9999
#     samples = sorted(samples, key=sort_key)

#     doc = Document()

#     # ===== Tabel 1 =====
#     p = doc.add_paragraph("Tabel 1. Samenstelling analysemonsters.")
#     for r in p.runs:
#         r.font.size = Pt(9); r.italic = True

#     cols1 = [MC, SAM, "Boornummer\n(traject in m - mv.)", OND]
#     t1 = doc.add_table(rows=1, cols=len(cols1))
#     t1.style = "Table Grid"
#     hdr1 = t1.rows[0].cells
#     for j, name in enumerate(cols1):
#         # nette multi-line header
#         hdr1[j].text = ""
#         for (k, line) in enumerate(str(name).split("\n")):
#             (hdr1[j].paragraphs[0] if k == 0 else hdr1[j].add_paragraph()).add_run(line).bold = True
#         for run in hdr1[j].paragraphs[0].runs:
#             run.font.size = Pt(9)

#     for s in samples:
#         row = t1.add_row().cells
#         row[0].text = s.get(MC, "")
#         row[1].text = s.get(SAM, "")
#         # multi-line cell via paragrafen ipv '\n'
#         row[2].text = ""
#         lines = s.get(BN, []) or []
#         if not isinstance(lines, list): lines = [str(lines)]
#         for idx, line in enumerate(lines):
#             (row[2].paragraphs[0] if idx == 0 else row[2].add_paragraph()).add_run(str(line))
#         row[3].text = s.get(OND, "")
#         # font-size uniform
#         for c in row:
#             for p in c.paragraphs:
#                 for r in p.runs:
#                     r.font.size = Pt(9)

#     # ===== Tabel 2 =====
#     p = doc.add_paragraph("Tabel 2. Samenvatting toetsing milieuhygiënische kwaliteit grond.")
#     for r in p.runs:
#         r.font.size = Pt(9); r.italic = True

#     cols2 = [MC, SAM, "Boornummer\n(traject in m - mv.)", SKF, KKA]
#     t2 = doc.add_table(rows=1, cols=len(cols2))
#     t2.style = "Table Grid"
#     hdr2 = t2.rows[0].cells
#     for j, name in enumerate(cols2):
#         hdr2[j].text = ""
#         for (k, line) in enumerate(str(name).split("\n")):
#             (hdr2[j].paragraphs[0] if k == 0 else hdr2[j].add_paragraph()).add_run(line).bold = True
#         for run in hdr2[j].paragraphs[0].runs:
#             run.font.size = Pt(9)

#     # legenda-detectie (veilig, zonder regex-finesses)
#     tokens = set()
#     for s in samples:
#         row = t2.add_row().cells
#         row[0].text = s.get(MC, "")
#         row[1].text = s.get(SAM, "")
#         row[2].text = ""
#         lines = s.get(BN, []) or []
#         if not isinstance(lines, list): lines = [str(lines)]
#         for idx, line in enumerate(lines):
#             (row[2].paragraphs[0] if idx == 0 else row[2].add_paragraph()).add_run(str(line))
#         skf_val = s.get(SKF, "") or ""
#         row[3].text = skf_val
#         row[4].text = s.get(KKA, "") or ""
#         for c in row:
#             for p in c.paragraphs:
#                 for r in p.runs:
#                     r.font.size = Pt(9)
#         for tok in ("L/N","W","IND","I","MV","SV"):
#             if tok in skf_val.upper(): tokens.add("IND" if tok=="I" else tok)

#     # legenda (simpel)
#     only_ln = bool(tokens) and tokens <= {"L/N"}
#     doc.add_paragraph(
#         "L/N : geen verontreinigingen aangetoond (de waarden overschrijden de kwaliteitseis voor klasse 'landbouw / natuur' niet)"
#         if only_ln else
#         "L/N: geen verontreinigingen aangetoond ...\nW: wonen ...\nIND: industrie ...\nMV: matig verontreinigd ...\nSV: sterk verontreinigd ..."
#     ).runs[0].font.size = Pt(8)

#     bio = io.BytesIO()
#     doc.save(bio)
#     bio.seek(0)
#     return bio


# --- DROP-IN REPLACER ---
import io
import re
from typing import List, Dict
from docx import Document
from docx.shared import Pt, RGBColor, Cm
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

_L_N_LEGEND = "L/N : geen verontreinigingen aangetoond (de waarden overschrijden de kwaliteitseis voor klasse 'landbouw / natuur' niet)"
_W_LEGEND = "W : wonen (licht verontreinigd; de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'wonen')"
_IND_LEGEND = "IND : industrie (licht verontreinigd; de aangetoonde waarden voldoen aan de kwaliteitseis van klasse 'industrie')"
_MV_LEGEND = "MV : matig verontreinigd (de aangetoonde waarden voldoen aan de kwaliteitseis voor klasse 'matig verontreinigd')"
_SV_LEGEND = "SV : sterk verontreinigd (de aangetoonde waarden overschrijden de norm behorend bij de kwaliteitseis voor klasse 'matig verontreinigd' / interventiewaarde bodemkwaliteit (I))"

_LEGEND_IND = "IND"
_LEGEND_W = "W"
_LEGEND_MV = "MV"
_LEGEND_SV = "SV"

_NEN_5740_LEGEND = (
        "NEN 5740 grond:\t\tmetalen (barium, cadmium, kobalt, koper, kwik, lood, molybdeen, nikkel, zink), PAK (polycyclische"
        "\n\t\t\taromatische koolwaterstoffen), PCB (polychloorbifenylen), minerale olie, droge stof-, lutum- en"
        "\n\t\t\torganische stofgehalte."
    )

_PFAS_LEGEND = "PFAS:\t\t\tper- en polyfluoralkylverbindingen"

_NEN_5740_LEGEND_FONT_SIZE = 8
_PFAS_LEGEND_FONT_SIZE = 8

_MM_LEGEND_FONT_SIZE = 8
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

def _set_table_width(table, total_cm: float):
    tbl = table._tbl
    tblPr = tbl.tblPr
    for el in tblPr.findall(qn('w:tblW')):
        tblPr.remove(el)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:type'), 'dxa')
    tblW.set(qn('w:w'), str(_dxa_from_cm(total_cm)))
    tblPr.append(tblW)

def _set_col_widths_tcW(table, widths_cm):
    """Zet per cel een expliciete tcW (type=dxa). Dit is Windows-safe en respecteert kolombreedtes."""
    table.autofit = False
    widths_tw = [_dxa_from_cm(w) for w in widths_cm]
    for col_idx, tw in enumerate(widths_tw):
        for row in table.rows:
            cell = row.cells[col_idx]
            tcPr = cell._tc.get_or_add_tcPr()
            for el in tcPr.findall(qn('w:tcW')):
                tcPr.remove(el)
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:type'), 'dxa')
            tcW.set(qn('w:w'), str(tw))
            tcPr.append(tcW)

def _borders_horizontal_only(table):
    """Alleen top/bottom/insideH, geen insideV/left/right (Windows-safe)."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    for el in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(el)
    borders = OxmlElement('w:tblBorders')
    # top/bottom stevig
    def add(tag, val='single', sz='8', color='000000'):
        el = OxmlElement(tag); el.set(qn('w:val'), val); el.set(qn('w:sz'), sz); el.set(qn('w:color'), color); borders.append(el)
    add('w:top'); add('w:bottom')
    # insideH dunner
    insideH = OxmlElement('w:insideH')
    insideH.set(qn('w:val'), 'single'); insideH.set(qn('w:sz'), '4'); insideH.set(qn('w:color'), '000000')
    borders.append(insideH)
    # disable verticale & zijkanten
    for tag in ('w:insideV','w:left','w:right'):
        el = OxmlElement(tag); el.set(qn('w:val'), 'nil'); borders.append(el)
    tblPr.append(borders)

# ---------- exporter ----------

def export_to_docx(samples: List[Dict]) -> io.BytesIO:
    # sortering op monstercode
    def sort_key(s):
        try:    return int(''.join(c for c in s.get(MC,'') if c.isdigit()))
        except: return 9999
    samples = sorted(samples, key=sort_key)

    doc = Document()
    # style "Normal" richting Calibri 9pt
    try:
        normal = doc.styles['Normal']
        normal.font.name = 'Calibri'; normal.font.size = Pt(9)
        n = normal._element; rPr = n.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts')) or OxmlElement('w:rFonts')
        if rFonts.getparent() is None: rPr.append(rFonts)
        for k in ('w:ascii','w:hAnsi','w:cs'): rFonts.set(qn(k), 'Calibri')
    except Exception: pass

    # ===== Tabel 1 =====
    p1 = doc.add_paragraph("Tabel 1. Samenstelling analysemonsters.")
    for r in p1.runs: _force_calibri(r, size_pt=9, italic=True)

    cols1 = [MC, SAM, "Boornummer\n(traject in m - mv.)", OND]
    widths_t1 = [2.13, 4.50, 3.75, 4.87]  # cm  (som ≈ 15.25)
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
        if not isinstance(lines, list): lines = [str(lines)]
        for i, line in enumerate(lines):
            (_p := (row[2].paragraphs[0] if i == 0 else row[2].add_paragraph())).add_run(str(line))
        row[3].text = s.get(OND, "")
        for c in row:
            _cell_runs_calibri(c, size_pt=9)

    _set_table_width(t1, sum(widths_t1))
    _set_col_widths_tcW(t1, widths_t1)
    _borders_horizontal_only(t1)

    # notities
    n1 = doc.add_paragraph("MM = mengmonster")
    [_force_calibri(r, size_pt=_MM_LEGEND_FONT_SIZE) for r in n1.runs]
    n2 = doc.add_paragraph(_NEN_5740_LEGEND)
    [_force_calibri(r, size_pt=_NEN_5740_LEGEND_FONT_SIZE) for r in n2.runs]
    n3 = doc.add_paragraph(_PFAS_LEGEND)
    [_force_calibri(r, size_pt=_PFAS_LEGEND_FONT_SIZE) for r in n3.runs]

    # ===== Tabel 2 =====
    p2 = doc.add_paragraph("Tabel 2. Samenvatting toetsing milieuhygiënische kwaliteit grond.")
    for r in p2.runs: _force_calibri(r, size_pt=9, italic=True)

    cols2 = [MC, SAM, "Boornummer\n(traject in m - mv.)", SKF, KKA]
    widths_t2 = [2.13, 2.75, 3.75, 3.00, 3.50]  # cm (som ≈ 15.13)
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
        if not isinstance(lines, list): lines = [str(lines)]
        for i, line in enumerate(lines):
            (_p := (row[2].paragraphs[0] if i == 0 else row[2].add_paragraph())).add_run(str(line))
        skf_val = (s.get(SKF, "") or "")
        row[3].text = skf_val
        row[4].text = s.get(KKA, "") or ""
        for c in row: _cell_runs_calibri(c, size_pt=9)
        for tok in _CLASS_TOKEN_RE.findall(skf_val.upper()):
            tokens.add("IND" if tok == "I" else tok)

    _set_table_width(t2, sum(widths_t2))
    _set_col_widths_tcW(t2, widths_t2)
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
