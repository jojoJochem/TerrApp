import pytest

from parser import (
    GW_EGV,
    GW_GWS,
    GW_PB,
    MC,
    BN,
    OND,
    parse_excel_file,
    parse_excel_to_samples,
)


def _codes(samples):
    return [s[MC] for s in samples]


def test_parser_returns_samples_and_project_code(make_tabel_workbook, tmp_xlsm_path):
    xlsm_bytes = make_tabel_workbook(
        project_code="T.25.16109",
        include_pfas=True,
        samples=("MM01", "02"),
    )
    tmp_xlsm_path.write_bytes(xlsm_bytes)

    samples, project_code = parse_excel_to_samples(str(tmp_xlsm_path))

    assert project_code == "T.25.16109"
    assert isinstance(samples, list)
    assert len(samples) == 2

    codes = set(_codes(samples))
    assert "MM01" in codes
    assert "02" in codes or "2" in codes

    s0 = samples[0]
    boors = s0[BN]

    assert any("01 t/m 03" in line for line in boors)
    assert any("05" in line for line in boors)

    assert "NEN 5740" in s0[OND]
    assert "arseen" in s0[OND]
    assert "PFAS" in s0[OND]


@pytest.mark.parametrize(
    ("filename", "expected_codes", "expected_groundwater_count"),
    [
        (
            "24-08-08_bijlage_2.1_1_grond_T2413196.xlsm",
            {"01", "MM01"},
            0,
        ),
        (
            "25-06-03_bijlage_2_1_grond_T2513603.xlsm",
            {"MM01", "MM02", "MM03"},
            0,
        ),
        (
            "25-06-03_bijlage_2_2_grond_T2513603.xlsm",
            {"MM04", "MM05", "MM06"},
            0,
        ),
        (
            "2.1 A02 A02 A04.xlsm",
            {"A02", "A04"},
            0,
        ),
        (
            "2.1 MMA01 MMA02 MMA03.xlsm",
            {"MMA01", "MMA02", "MMA03"},
            0,
        ),
        (
            "2.1 D04 D05 D06.xlsm",
            {"D04", "D05", "D06"},
            0,
        ),
        (
            "2.1 MMA04 MMA05 MMA06.xlsm",
            {"MMA04", "MMA05", "MMA06"},
            0,
        ),
        (
            "2.2 A04 B09 D02 D06.xlsm",
            set(),
            4,
        ),
    ],
)
def test_parser_real_uploaded_files(fixture_file, filename, expected_codes, expected_groundwater_count):
    path = fixture_file(filename)

    ground_samples, groundwater_samples, project_code = parse_excel_file(str(path))

    assert isinstance(project_code, str)

    if expected_codes:
        assert len(ground_samples) >= len(expected_codes)

        actual_codes = set(_codes(ground_samples))
        assert expected_codes <= actual_codes

        for sample in ground_samples:
            assert sample[MC]
            assert isinstance(sample[BN], list)
            assert sample[OND].startswith("NEN 5740")

    else:
        assert ground_samples == []

    assert len(groundwater_samples) == expected_groundwater_count

    for sample in groundwater_samples:
        assert sample[GW_PB]
        assert sample[GW_GWS]
        assert sample[GW_EGV]


def test_parser_real_groundwater_file_has_expected_values(fixture_file):
    path = fixture_file("2.2 A04 B09 D02 D06.xlsm")

    ground_samples, groundwater_samples, project_code = parse_excel_file(str(path))

    assert ground_samples == []
    assert len(groundwater_samples) == 4

    by_code = {s[GW_PB]: s for s in groundwater_samples}

    assert set(by_code) == {"A04", "B09", "D02", "D06"}

    assert by_code["D02"][GW_GWS] == "1,70"
    assert by_code["D06"][GW_GWS] == "1,40"

    assert by_code["D02"][GW_EGV] == "1.400"
    assert by_code["D06"][GW_EGV] == "1.100"


def test_parser_all_real_ground_files_together(real_ground_files):
    all_ground_samples = []
    all_groundwater_samples = []

    for path in real_ground_files:
        ground_samples, groundwater_samples, _project_code = parse_excel_file(str(path))
        all_ground_samples.extend(ground_samples)
        all_groundwater_samples.extend(groundwater_samples)

    # 2 + 3 + 3 + 3 + 3 + 3 + 3 = 20 grondmonsters
    assert len(all_ground_samples) == 20
    assert all_groundwater_samples == []

    codes = _codes(all_ground_samples)

    for expected in [
        "01",
        "MM01",
        "MM02",
        "MM03",
        "MM04",
        "MM05",
        "MM06",
        "A02",
        "A04",
        "MMA01",
        "MMA02",
        "MMA03",
        "D04",
        "D05",
        "D06",
        "MMA04",
        "MMA05",
        "MMA06",
    ]:
        assert expected in codes


def test_parser_all_real_files_together(real_excel_files):
    all_ground_samples = []
    all_groundwater_samples = []

    for path in real_excel_files:
        ground_samples, groundwater_samples, _project_code = parse_excel_file(str(path))
        all_ground_samples.extend(ground_samples)
        all_groundwater_samples.extend(groundwater_samples)

    assert len(all_ground_samples) == 20
    assert len(all_groundwater_samples) == 4

    assert {s[GW_PB] for s in all_groundwater_samples} == {"A04", "B09", "D02", "D06"}