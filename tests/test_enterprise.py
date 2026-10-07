"""Run against the real example files: set EXAMPLES to the folder holding them."""
import io, os, random, sys, zipfile
import pytest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import app

EX = os.environ.get("EXAMPLES", "/root/.claude/uploads/12d12e69-6887-5c44-990e-a999b131df49")

def find(part):
    for n in os.listdir(EX):
        if part in n:
            return open(os.path.join(EX, n), "rb").read()
    pytest.skip(f"{part} not available")

@pytest.fixture(scope="module")
def digs():
    a = app.parse_dig_package(find("AID_115"), "AID_115.xlsm")
    b = app.parse_dig_package(find("AID_116"), "AID_116.xlsm")
    return a, b

def test_counts(digs):
    assert len(digs[0]) == 59 and len(digs[1]) == 63

def test_dig_02b_fields(digs):
    d = next(x for x in digs[0] if x.number == "02B")
    assert (d.odometer, d.station_ft) == (57975.75, 332977)
    assert d.upstream_reference == "AGM 090" and d.downstream_reference == "AGM 100"
    assert round(d.us_weld_distance, 2) == 9.35 and round(d.ds_weld_distance, 2) == 48.02
    assert d.county_state == "Morris County, KS" and d.tract_number == "5-K-MO-30"
    assert d.nearest_reference_label == "AGM 100 1717.92'"

def test_kmz_matches_by_aid_and_feature(digs):
    k = app.parse_kmz(find("KMZ"), "x.kmz")
    assert len(k.lines) == 1 and len(k.digs) == 129
    alld = digs[0] + digs[1]
    app.reconcile_with_kmz(alld, k)
    assert not [w for d in alld for w in d.warnings]

def test_report_and_cheat_sheet(digs):
    import openpyxl
    tpl = find("Autopopulates")
    d = next(x for x in digs[1] if x.number == "05A")
    out, warn = app.build_staking_report(tpl, d, app.ReportSettings("Hayden Miller"),
                                         app.read_template_info(tpl).phones)
    assert not warn
    wb = openpyxl.load_workbook(io.BytesIO(out), data_only=True)
    sr = wb["StakingReport"]
    assert sr["G4"].value == "Dig # 05A" and sr["C12"].value == "580-585-7118"
    assert sr["J5"].value == d.station_ft and sr["C24"].value == d.hca
    z = zipfile.ZipFile(io.BytesIO(out))
    assert "xl/vbaProject.bin" in z.namelist()
    # pre-dig photo placeholders must keep their own media part
    assert "xl/media/image6.jpeg" in z.namelist()
    ch = app.build_cheat_sheet(find("for_ascending"), digs[0], auto_notes=True)
    ws = openpyxl.load_workbook(io.BytesIO(ch)).worksheets[0]
    assert ws["A6"].value.endswith("U/S Weld") and ws["B6"].value == "=B7-D6"
    first = min(digs[0], key=lambda x: x.odometer)
    assert (ws["I7"].value, ws["J7"].value) == (first.latitude, first.longitude)
    assert ws["I6"].value is None

def test_aerial_exact_span():
    z, f, mpp = app.aerial_geometry(38.7, 2600, 1600, 256, 19)
    assert abs(mpp * 1600 * 3.28084 - 2600) < 1
    assert 1.0 <= f < 2.0
