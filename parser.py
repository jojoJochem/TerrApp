# parser.py
import re
import pandas as pd
from typing import List, Dict, Any
from collections import OrderedDict

# Output-keys
MC = 'Monstercode'
SAM = 'Samenstelling'
BN = 'Boornummer'
OND = 'Onderzochte parameters'
SKF = 'Stofspecifieke kwaliteitsklassen'
KKA = 'Kwaliteitsklasse analysemonster'

# Ankers in sheet "Tabel"
ANCHOR_MM = 'Mengmonster / boring'
ANCHOR_MS = 'Monstersamenstelling'              # volgende regel: '(traject in m - mv.)'
ANCHOR_RBK = 'Kwaliteitsklasse Rbk'
ANCHOR_PFAS = 'Kwaliteitsklasse PFAS'
ANCHOR_TOT = 'Kwaliteitsklasse totaal'

# PFAS-rijen herkennen o.b.v. tekst in eerste kolom
PFAS_ROWS_KEYWORDS = ['PF', 'GenX', 'FOSA', 'FOSAA', 'diPAP', 'PFAS']

_MM_CODE = re.compile(r'^\s*MM\d+', re.IGNORECASE)
_INT_TOKEN = re.compile(r'\b\d+\b')
_BOOR_RE = re.compile(r'^\s*(\d+)\s*\(([^)]+)\)\s*$')  # "01 (0,50-1,00)" -> num="01", depth="0,50-1,00"


def _find_row(sheet: pd.DataFrame, startswith_text: str) -> int:
    for i in range(sheet.shape[0]):
        v = str(sheet.iat[i ,0]).strip()
        if v.startswith(startswith_text):
            return i
    raise ValueError(f"Anchor row not found for: {startswith_text}")


def _list_pfas_rows(sheet: pd.DataFrame) -> List[int]:
    idxs = []
    for i in range(sheet.shape[0]):
        v = sheet.iat[i,0]
        if isinstance(v,str) and any(k in v for k in PFAS_ROWS_KEYWORDS):
            if 'Kwaliteitsklasse' not in v:
                idxs.append(i)
    return sorted(set(idxs))


# def _get_sample_start_cols(sheet: pd.DataFrame, row_mm: int) -> List[int]:
#     cols = []
#     for j in range(sheet.shape[1]):
#         v = sheet.iat[row_mm, j]
#         if isinstance(v, str) and re.match(r'^MM\d+', v.strip()):
#             cols.append(j)
#     return cols


def _get_sample_start_cols(sheet: pd.DataFrame, row_mm: int) -> List[int]:
    """
    Vind kolommen in de 'Mengmonster / boring'-rij die een sample starten:
    - Mengmonster: tekst begint met 'MM' + cijfers (bv. MM01).
    - Losse boorkop(pen): er staat ≥1 geheel getal in de cel (bv. '01', '7-8', '12 13').
    """
    cols = []
    for j in range(sheet.shape[1]):
        v = sheet.iat[row_mm, j]
        if not isinstance(v, str):
            continue
        s = v.strip()
        if not s:
            continue

        # 1) Mengmonsters (MMxx)
        if _MM_CODE.match(s):
            cols.append(j)
            continue

        # 2) Boorkop(pen): minstens één integer aanwezig
        if len(_INT_TOKEN.findall(s)) >= 1:
            cols.append(j)

    return cols


def _fetch_class(sheet: pd.DataFrame, row: int, col_hint: int) -> str:
    # Sommige bestanden hebben 1 kolom offset; kijk in een klein venster
    for jj in [col_hint, col_hint+1, col_hint-1, col_hint+2]:
        if 0 <= jj < sheet.shape[1]:
            val = sheet.iat[row, jj]
            if isinstance(val, str) and val.strip():
                return val.strip()
    return ""


def _join_lines(*vals: Any) -> str:
    parts = []
    for v in vals:
        if isinstance(v, str) and v.strip().lower() != 'zz' and v.strip():
            parts.append(v.strip().strip(','))
    return ', '.join(parts)


def _compress_sequence(nums: List[int]) -> List[List[int]]:
    """Splits een gesorteerde lijst ints in opeenvolgende runs."""
    if not nums:
        return []
    runs = []
    start = nums[0]
    prev = nums[0]
    for x in nums[1:]:
        if x == prev + 1:
            prev = x
        else:
            runs.append([start, prev])
            start = prev = x
    runs.append([start, prev])
    return runs


def _group_and_format_boors(boors: List[str]) -> List[str]:
    """
    Neem boorregels als '01 (0,50-1,00)' en groepeer per diepte.
    - Reeksen van >=3 → '01 t/m 03'
    - Reeksen van 2 → '01, 02'
    - Enkele → '01'
    Retourneert een lijst regels: '01 t/m 03, 05, 08 t/m 11 (0,50-1,00)'
    """
    if not boors:
        return []

    # depth -> { 'nums': [int...], 'seen': [("01", 1), ...], 'order': index }
    groups = OrderedDict()
    order = 0
    for item in boors:
        m = _BOOR_RE.match(item)
        if not m:
            # onbekend formaat: zet op aparte key zodat niets verloren gaat
            groups.setdefault(item, {'nums': [], 'seen': [], 'order': order})
            order += 1
            continue
        num_str, depth = m.group(1), m.group(2)
        num_int = int(num_str)
        if depth not in groups:
            groups[depth] = {'nums': [], 'seen': [], 'order': order}
            order += 1
        groups[depth]['nums'].append(num_int)
        groups[depth]['seen'].append((num_str, num_int))

    # Maak outputregels per depth, in encounter-volgorde
    lines = []
    for depth, data in sorted(groups.items(), key=lambda kv: kv[1]['order']):
        nums = sorted(set(data['nums']))  # sorteer en dedup
        # breedte bepalen voor zero-padding obv daadwerkelijk geziene strings
        width = 0
        for ns, _ni in data['seen']:
            width = max(width, len(ns))
        width = max(width, 2)  # minimaal 2 is gebruikelijk ("01")

        runs = _compress_sequence(nums)

        parts = []
        for a, b in runs:
            if b - a + 1 >= 3:
                parts.append(str(a).zfill(width) + " t/m " + str(b).zfill(width))
            elif b - a + 1 == 2:
                parts.append(str(a).zfill(width))
                parts.append(str(b).zfill(width))
            else:
                parts.append(str(a).zfill(width))

        # Combineer en plak de depth erachter
        if data['nums']:  # normale depth-groep
            lines.append(f"{', '.join(parts)} ({depth})")
        else:
            # fallback voor niet-parsebare items
            lines.append(depth)

    return lines


def build_onderzochte_parameters(sheet: pd.DataFrame, sample_col: int) -> str:
    params = ['NEN 5740 grond']
    # Arseen aanwezig?
    r_as = None
    for i in range(sheet.shape[0]):
        if sheet.iat[i ,0] == 'Arseen':
            r_as = i
            break
    if r_as is not None:
        val = sheet.iat[r_as, sample_col]
        if isinstance(val, (int, float)) or (isinstance(val, str) and val.strip() not in ('', '--', 'NaN')):
            params.append('arseen')

    # PFAS aanwezig in een van de PFAS-rijen?
    has_pfas = False
    for rr in _list_pfas_rows(sheet):
        if sample_col < sheet.shape[1]:
            v = sheet.iat[rr, sample_col]
            if isinstance(v,(int, float)) or (isinstance(v, str) and v.strip() not in ('', '--', 'NaN')):
                has_pfas = True
                break
    if has_pfas:
        params.append('PFAS')

    return ', '.join(params)


# def parse_excel_to_samples(path: str) -> List[Dict[str, Any]]:
#     # Lees het sheet "Tabel" (in jouw bestanden staat dáár de info voor boornummers & klassen)
#     sheet = pd.read_excel(path, sheet_name='Tabel', header=None, engine='openpyxl')

#     row_mm = _find_row(sheet, ANCHOR_MM)
#     row_ms = _find_row(sheet, ANCHOR_MS)
#     row_rb = _find_row(sheet, ANCHOR_RBK)
#     row_pf = _find_row(sheet, ANCHOR_PFAS)
#     row_tt = _find_row(sheet, ANCHOR_TOT)

#     start_cols = _get_sample_start_cols(sheet, row_mm)

#     samples = []
#     for c in start_cols:
#         raw_code = str(sheet.iat[row_mm, c]).strip()

#         # Monstercode bepalen
#         if _MM_CODE.match(raw_code):
#             code = raw_code
#         else:
#             ints = _INT_TOKEN.findall(raw_code)
#             if len(ints) >= 1:
#                 # Voor boorkoppen de integers als code nemen (bv. '7-8' -> '7 8', '01' -> '01')
#                 code = " ".join(ints)
#             else:
#                 # geen bruikbare code → sla kolom over
#                 continue

#         samenstelling = _join_lines(sheet.iat[row_mm+1, c], sheet.iat[row_mm+2, c], sheet.iat[row_mm+3, c])

#         # boornummers: rijen onder 'Monstersamenstelling'
#         boors = []
#         for rr in range(row_ms, row_ms+12):
#             if rr >= sheet.shape[0]: break
#             val = sheet.iat[rr, c]
#             if isinstance(val, str):
#                 s = val.strip()
#                 if s.lower() != 'zz' and re.match(r'^\d+\s*\(.*\)$', s):
#                     boors.append(s)

#         klas_rb = _fetch_class(sheet, row_rb, c)
#         klas_pf = _fetch_class(sheet, row_pf, c)
#         klas_tot = _fetch_class(sheet, row_tt, c)

#         if 'landbouw' in (klas_rb or '').lower() and 'landbouw' in (klas_pf or '').lower():
#             stofspec = 'alle: L/N'
#         else:
#             def abbr(k: str) -> str:
#                 if not isinstance(k, str):
#                     return 'n/b'
#                 kl = k.lower()
#                 if 'landbouw' in kl:
#                     return 'L/N'
#                 if 'wonen' in kl:
#                     return 'W'
#                 if 'industrie' in kl:
#                     return 'I'
#                 return k
#             stofspec = f"RbK: {abbr(klas_rb)}, PFAS: {abbr(klas_pf)}"

#         params = build_onderzochte_parameters(sheet, c)

#         samples.append({
#             MC: code,
#             SAM: samenstelling,
#             BN: boors,
#             OND: params,
#             SKF: stofspec,
#             KKA: klas_tot
#         })

#     return samples


def parse_excel_to_samples(path: str) -> List[Dict[str, Any]]:
    """
    Leest het sheet 'Tabel' en bouwt per kolom (MM of losse boorkop) een sample-record
    voor de Word-tabellen. Boornummers worden per diepte gegroepeerd en reeksen samengevouwen.
    Vereist hulpfuncties/constanten:
      - _find_row, _get_sample_start_cols, _fetch_class, _join_lines, build_onderzochte_parameters
      - _group_and_format_boors (groepeert en vouwt boor-reeksen)
      - regex _MM_CODE en _INT_TOKEN
      - output keys: MC, SAM, BN, OND, SKF, KKA
    """
    sheet = pd.read_excel(path, sheet_name='Tabel', header=None, engine='openpyxl')

    row_mm = _find_row(sheet, ANCHOR_MM)
    row_ms = _find_row(sheet, ANCHOR_MS)
    row_rb = _find_row(sheet, ANCHOR_RBK)
    row_pf = _find_row(sheet, ANCHOR_PFAS)
    row_tt = _find_row(sheet, ANCHOR_TOT)

    start_cols = _get_sample_start_cols(sheet, row_mm)

    samples = []
    for c in start_cols:
        raw_code = str(sheet.iat[row_mm, c]).strip()

        # Monstercode bepalen: MMxx direct overnemen, anders integers uit de kop gebruiken
        if _MM_CODE.match(raw_code):
            code = raw_code
        else:
            ints = _INT_TOKEN.findall(raw_code)  # alleen gehele getallen (leading zeros behouden via zfill later)
            if len(ints) >= 1:
                code = " ".join(ints)  # bv. "7-8" -> "7 8", "01" -> "01"
            else:
                # Geen bruikbare code → kolom overslaan
                continue

        # Samenstelling komt uit de 3 regels direct onder de kop
        samenstelling = _join_lines(
            sheet.iat[row_mm + 1, c],
            sheet.iat[row_mm + 2, c],
            sheet.iat[row_mm + 3, c]
        )

        # Boornummers verzamelen uit blok onder 'Monstersamenstelling'
        boors_raw = []
        for rr in range(row_ms, row_ms + 12):
            if rr >= sheet.shape[0]:
                break
            val = sheet.iat[rr, c]
            if isinstance(val, str):
                s = val.strip()
                # verwacht formaat: "<int> (<diepte>)", ignore 'zz'
                if s and s.lower() != 'zz' and re.match(r'^\d+\s*\(.*\)$', s):
                    boors_raw.append(s)

        # Groepeer boornummers per diepte en vouw reeksen samen (01 t/m 03, 05, 08 t/m 11)
        boors_grouped = _group_and_format_boors(boors_raw)

        # Klassen ophalen (met kleine tolerantie in _fetch_class)
        klas_rb = _fetch_class(sheet, row_rb, c)
        klas_pf = _fetch_class(sheet, row_pf, c)
        klas_tot = _fetch_class(sheet, row_tt, c)

        # Stofspecifieke kwaliteitsklassen opmaken
        if 'landbouw' in (klas_rb or '').lower() and 'landbouw' in (klas_pf or '').lower():
            stofspec = 'alle: L/N'
        else:
            def abbr(k: str) -> str:
                if not isinstance(k, str):
                    return 'n/b'
                kl = k.lower()
                if 'landbouw' in kl:
                    return 'L/N'
                if 'wonen' in kl:
                    return 'W'
                if 'industrie' in kl:
                    return 'I'
                return k
            stofspec = f"PFOS: {abbr(klas_pf)}, Overig: {abbr(klas_rb)}"

        # Onderzochte parameters bepalen
        params = build_onderzochte_parameters(sheet, c)

        samples.append({
            MC: code,
            SAM: samenstelling,
            BN: boors_grouped,   # gegroepeerde boorregels (meerdere regels bij meerdere dieptes)
            OND: params,
            SKF: stofspec,
            KKA: klas_tot
        })

    return samples
