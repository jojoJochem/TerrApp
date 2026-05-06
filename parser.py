import os
import re
from collections import OrderedDict
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

# Output-keys grond
MC = "Monstercode"
SAM = "Samenstelling"
BN = "Boornummer"
OND = "Onderzochte parameters"
SKF = "Stofspecifieke kwaliteitsklassen"
KKA = "Kwaliteitsklasse analysemonster"

# Output-keys grondwater
GW_PB = "Peilbuis"
GW_FS = "Filterstelling"
GW_GWS = "Grondwaterstand"
GW_PH = "pH"
GW_EGV = "EGV"
GW_TR = "Troebelheid"

# Ankers grond
ANCHOR_MM = "Mengmonster / boring"
ANCHOR_MS = "Monstersamenstelling"
ANCHOR_RBK = "Kwaliteitsklasse Rbk"
ANCHOR_PFAS = "Kwaliteitsklasse PFAS"
ANCHOR_TOT = "Kwaliteitsklasse totaal"

# Ankers grondwater
ANCHOR_GW_PB = "Peilbuis"

PFAS_ROWS_KEYWORDS = ["PF", "GenX", "FOSA", "FOSAA", "diPAP", "PFAS"]

_PROJECT_CODE_RE = re.compile(r"T\.\s*\d{2}\.\s*\d{3,6}")
_CODE_SPLIT_RE = re.compile(r"[A-Za-z]*\d+[A-Za-z0-9]*")

# Ondersteunt: 01, MM01, MMA01, A02, D06, B09, 7-8.
_SAMPLE_CODE_RE = re.compile(
    r"^\s*(?:MM[A-Za-z]*\d+[A-Za-z0-9]*|[A-Za-z]*\d+[A-Za-z0-9]*(?:\s*[-–]\s*[A-Za-z]*\d+[A-Za-z0-9]*)*)\s*$",
    re.IGNORECASE,
)

# Ondersteunt: 01 (0,00-0,50), A01 (0,00-0,50), D06 (1,50-2,00).
_BOOR_RE = re.compile(
    r"^\s*([A-Za-z]*\d+[A-Za-z0-9]*)\s*\(([^)]+)\)\s*$",
    re.IGNORECASE,
)


def _read_sheet(path: str) -> List[List[Any]]:
    ext = os.path.splitext(path)[1].lower()

    if ext in {".xlsx", ".xlsm"}:
        df = pd.read_excel(path, sheet_name="Tabel", header=None, engine="openpyxl")
    else:
        df = pd.read_excel(path, sheet_name="Tabel", header=None)

    df = df.where(pd.notna(df), None)
    return df.values.tolist()


def _get_cell(sheet: List[List[Any]], row: int, col: int) -> Any:
    if row < 0 or row >= len(sheet):
        return None
    if col < 0 or col >= len(sheet[row]):
        return None
    return sheet[row][col]


def _row_len(sheet: List[List[Any]]) -> int:
    return max((len(r) for r in sheet), default=0)


def _as_text(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _clean_decimal_text(v: Any) -> str:
    s = _as_text(v)

    if not s or s.lower() in {"nan", "--", "zz"}:
        return ""

    try:
        f = float(s.replace(",", "."))
        if abs(f - round(f)) < 1e-9:
            return str(int(round(f)))
        out = f"{f:.2f}".rstrip("0").rstrip(".")
        return out.replace(".", ",")
    except Exception:
        return s.replace(".", ",") if re.fullmatch(r"\d+\.\d+", s) else s


def _clean_egv_text(v: Any) -> str:
    s = _as_text(v)

    if not s or s.lower() in {"nan", "--", "zz"}:
        return ""

    try:
        n = int(float(s.replace(",", ".")))
        return f"{n:,}".replace(",", ".")
    except Exception:
        return _clean_decimal_text(v)


def _find_project_code(sheet: List[List[Any]]) -> str:
    for row in sheet:
        for v in row:
            m = _PROJECT_CODE_RE.search(_as_text(v))
            if m:
                return re.sub(r"\s+", "", m.group(0))
    return ""


def _find_row(sheet: List[List[Any]], startswith_text: str) -> int:
    for i, row in enumerate(sheet):
        first = _as_text(row[0] if row else "")
        if first.startswith(startswith_text):
            return i

    raise ValueError(f"Anchor row not found for: {startswith_text}")


def _find_row_optional(sheet: List[List[Any]], startswith_text: str) -> Optional[int]:
    try:
        return _find_row(sheet, startswith_text)
    except ValueError:
        return None


def _find_value_row(sheet: List[List[Any]], label: str) -> Optional[int]:
    label_l = label.lower()

    for i, row in enumerate(sheet):
        first = _as_text(row[0] if row else "").lower()
        if first.startswith(label_l):
            return i

    return None


def _find_row_exact_first_col(sheet: List[List[Any]], text: str) -> Optional[int]:
    for i, row in enumerate(sheet):
        if _as_text(row[0] if row else "") == text:
            return i
    return None


def _list_pfas_rows(sheet: List[List[Any]]) -> List[int]:
    idxs = []

    for i, row in enumerate(sheet):
        first = _as_text(row[0] if row else "")

        if not first:
            continue

        if any(k.lower() in first.lower() for k in PFAS_ROWS_KEYWORDS):
            if "kwaliteitsklasse" not in first.lower():
                idxs.append(i)

    return sorted(set(idxs))


def _get_sample_start_cols(sheet: List[List[Any]], row_mm: int) -> List[int]:
    cols = []

    for j in range(_row_len(sheet)):
        s = _as_text(_get_cell(sheet, row_mm, j))

        if not s:
            continue

        if s.lower().startswith("mengmonster"):
            continue

        if _SAMPLE_CODE_RE.match(s):
            cols.append(j)

    return cols


def _fetch_class(sheet: List[List[Any]], row: int, col_hint: int) -> str:
    for jj in [col_hint, col_hint + 1, col_hint - 1, col_hint + 2, col_hint + 3, col_hint + 4]:
        s = _as_text(_get_cell(sheet, row, jj))

        if s and s.lower() not in {"zz", "nan", "--"}:
            return s

    return ""


def _join_lines(*vals: Any) -> str:
    parts = []

    for v in vals:
        s = _as_text(v)

        if not s:
            continue

        if s.lower() in {"zz", "nan", "0", "--"}:
            continue

        parts.append(s.strip(","))

    return ", ".join(parts)


def _compress_sequence(nums: List[int]) -> List[List[int]]:
    if not nums:
        return []

    nums = sorted(nums)
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


def _split_prefix_number(code: str) -> Tuple[str, Optional[int], int]:
    m = re.match(r"^([A-Za-z]*)(\d+)([A-Za-z0-9]*)$", code)

    if not m:
        return code, None, 0

    prefix = m.group(1).upper()
    num = int(m.group(2))
    width = len(m.group(2))
    suffix = m.group(3)

    if suffix:
        return code, None, 0

    return prefix, num, width


def _format_code_run(prefix: str, a: int, b: int, width: int) -> List[str]:
    aa = f"{prefix}{str(a).zfill(width)}"
    bb = f"{prefix}{str(b).zfill(width)}"

    if b - a + 1 >= 3:
        return [f"{aa} t/m {bb}"]

    if b - a + 1 == 2:
        return [aa, bb]

    return [aa]


def _group_and_format_boors(boors: List[str]) -> List[str]:
    if not boors:
        return []

    depth_groups: "OrderedDict[str, List[str]]" = OrderedDict()
    fallback: List[str] = []

    for item in boors:
        s = _as_text(item)
        m = _BOOR_RE.match(s)

        if not m:
            fallback.append(s)
            continue

        code, depth = m.group(1), m.group(2)
        depth_groups.setdefault(depth, [])

        if code not in depth_groups[depth]:
            depth_groups[depth].append(code)

    lines: List[str] = []

    for depth, codes in depth_groups.items():
        grouped: "OrderedDict[str, List[Tuple[int, int, str]]]" = OrderedDict()
        loose: List[str] = []

        for code in codes:
            prefix, num, width = _split_prefix_number(code)

            if num is None:
                loose.append(code)
            else:
                grouped.setdefault(prefix, [])
                grouped[prefix].append((num, width, code))

        parts: List[str] = []

        for prefix, items in grouped.items():
            nums = sorted(set(num for num, _width, _code in items))
            width = max(max(width for _num, width, _code in items), 2)

            for a, b in _compress_sequence(nums):
                parts.extend(_format_code_run(prefix, a, b, width))

        parts.extend(loose)
        lines.append(f"{', '.join(parts)} ({depth})")

    lines.extend(fallback)
    return lines


def _value_present(v: Any) -> bool:
    s = _as_text(v)
    return bool(s and s.lower() not in {"zz", "nan", "--"})


def build_onderzochte_parameters(sheet: List[List[Any]], sample_col: int) -> str:
    params = ["NEN 5740 grond"]

    r_as = _find_row_exact_first_col(sheet, "Arseen")

    if r_as is not None:
        for jj in [sample_col, sample_col + 1, sample_col + 2, sample_col + 3, sample_col + 4]:
            if _value_present(_get_cell(sheet, r_as, jj)):
                params.append("arseen")
                break

    has_pfas = False

    for rr in _list_pfas_rows(sheet):
        for jj in [sample_col, sample_col + 1, sample_col + 2, sample_col + 3, sample_col + 4]:
            if _value_present(_get_cell(sheet, rr, jj)):
                has_pfas = True
                break

        if has_pfas:
            break

    if has_pfas:
        params.append("PFAS")

    return ", ".join(params)


def _abbr_class(k: str) -> str:
    kl = _as_text(k).lower()

    if not kl:
        return "n/b"

    if "landbouw" in kl or kl in {"ln", "l/n"}:
        return "L/N"

    if "wonen" in kl:
        return "W"

    if "industrie" in kl or kl == "ind":
        return "I"

    if "matig" in kl:
        return "MV"

    if "sterk" in kl:
        return "SV"

    return _as_text(k)


def _normalise_sample_code(raw_code: Any) -> str:
    raw_code = _as_text(raw_code)

    if not raw_code:
        return ""

    if _SAMPLE_CODE_RE.match(raw_code):
        return re.sub(r"\s+", " ", raw_code).strip()

    codes = _CODE_SPLIT_RE.findall(raw_code)
    return " ".join(codes)


def _collect_boors_for_sample(sheet: List[List[Any]], row_ms: int, sample_col: int) -> List[str]:
    boors_raw: List[str] = []

    for rr in range(row_ms, min(row_ms + 25, len(sheet))):
        first = _as_text(_get_cell(sheet, rr, 0)).lower()

        if first in {"droge stof (gew.%)"} or first.startswith("metalen"):
            break

        for cc in [sample_col, sample_col + 4]:
            s = _as_text(_get_cell(sheet, rr, cc))

            if not s:
                continue

            if s.lower() == "zz":
                continue

            if _BOOR_RE.match(s) and s not in boors_raw:
                boors_raw.append(s)

    return boors_raw


def _build_stofspecifieke_klasse(klas_rb: str, klas_pf: str, params: str) -> str:
    rb = _as_text(klas_rb)
    pf = _as_text(klas_pf)

    has_pfas = "PFAS" in params

    if has_pfas and "landbouw" in rb.lower() and "landbouw" in pf.lower():
        return "alle: L/N"

    if has_pfas:
        return f"PFOS: {_abbr_class(pf)}, Overig: {_abbr_class(rb)}"

    return f"Overig: {_abbr_class(rb)}"


def parse_ground_samples_from_sheet(sheet: List[List[Any]]) -> List[Dict[str, Any]]:
    row_mm = _find_row(sheet, ANCHOR_MM)
    row_ms = _find_row(sheet, ANCHOR_MS)
    row_rb = _find_row(sheet, ANCHOR_RBK)
    row_pf = _find_row(sheet, ANCHOR_PFAS)
    row_tt = _find_row(sheet, ANCHOR_TOT)

    start_cols = _get_sample_start_cols(sheet, row_mm)

    if not start_cols:
        raise ValueError("Geen analysemonsterkolommen gevonden.")

    samples: List[Dict[str, Any]] = []

    for c in start_cols:
        code = _normalise_sample_code(_get_cell(sheet, row_mm, c))

        if not code:
            continue

        samenstelling = _join_lines(
            _get_cell(sheet, row_mm + 1, c),
            _get_cell(sheet, row_mm + 2, c),
            _get_cell(sheet, row_mm + 3, c),
        )

        boors_grouped = _group_and_format_boors(
            _collect_boors_for_sample(sheet, row_ms, c)
        )

        klas_rb = _fetch_class(sheet, row_rb, c)
        klas_pf = _fetch_class(sheet, row_pf, c)
        klas_tot = _fetch_class(sheet, row_tt, c)

        if not klas_tot:
            klas_tot = klas_rb

        params = build_onderzochte_parameters(sheet, c)
        stofspec = _build_stofspecifieke_klasse(klas_rb, klas_pf, params)

        samples.append(
            {
                MC: code,
                SAM: samenstelling,
                BN: boors_grouped,
                OND: params,
                SKF: stofspec,
                KKA: klas_tot,
            }
        )

    return samples


def _get_groundwater_cols(sheet: List[List[Any]], row_pb: int) -> List[int]:
    cols = []

    for j in range(_row_len(sheet)):
        s = _as_text(_get_cell(sheet, row_pb, j))

        if not s:
            continue

        if s.lower() == "peilbuis":
            continue

        if _SAMPLE_CODE_RE.match(s):
            cols.append(j)

    return cols


def parse_groundwater_samples_from_sheet(sheet: List[List[Any]]) -> List[Dict[str, Any]]:
    row_pb = _find_row(sheet, ANCHOR_GW_PB)

    row_filter = _find_value_row(sheet, "Filterstelling")
    row_stand = _find_value_row(sheet, "Grondwaterstand")
    row_ph = _find_value_row(sheet, "pH")
    row_egv = _find_value_row(sheet, "Geleidbaarheid")
    row_tr = _find_value_row(sheet, "Troebelheid")

    required = [row_filter, row_stand, row_ph, row_egv, row_tr]

    if any(r is None for r in required):
        raise ValueError("Geen complete grondwater-veldmetingentabel gevonden.")

    rows: List[Dict[str, Any]] = []

    for c in _get_groundwater_cols(sheet, row_pb):
        peilbuis = _as_text(_get_cell(sheet, row_pb, c))

        if not peilbuis:
            continue

        filterstelling = _as_text(_get_cell(sheet, row_filter, c))
        grondwaterstand = _clean_decimal_text(_get_cell(sheet, row_stand, c))
        ph = _clean_decimal_text(_get_cell(sheet, row_ph, c))
        egv = _clean_egv_text(_get_cell(sheet, row_egv, c))
        troebelheid = _clean_decimal_text(_get_cell(sheet, row_tr, c))

        if not any([filterstelling, grondwaterstand, ph, egv, troebelheid]):
            continue

        rows.append(
            {
                GW_PB: peilbuis,
                GW_FS: filterstelling.replace(".", ","),
                GW_GWS: grondwaterstand,
                GW_PH: ph,
                GW_EGV: egv,
                GW_TR: troebelheid,
            }
        )

    return rows


def parse_excel_file(path: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], str]:
    sheet = _read_sheet(path)
    project_code = _find_project_code(sheet)

    ground_samples: List[Dict[str, Any]] = []
    groundwater_samples: List[Dict[str, Any]] = []

    if _find_row_optional(sheet, ANCHOR_MM) is not None:
        ground_samples = parse_ground_samples_from_sheet(sheet)

    if _find_row_optional(sheet, ANCHOR_GW_PB) is not None:
        groundwater_samples = parse_groundwater_samples_from_sheet(sheet)

    if not ground_samples and not groundwater_samples:
        raise ValueError("Geen grond- of grondwatertabel gevonden.")

    return ground_samples, groundwater_samples, project_code


def parse_excel_to_samples(path: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    Backwards compatible functie.
    Geeft alleen grondmonsters terug.
    """
    ground_samples, _groundwater_samples, project_code = parse_excel_file(path)
    return ground_samples, project_code