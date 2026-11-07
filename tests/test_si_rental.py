import io

import app as app_module
import csv
import datetime
import openpyxl
import pytest
import tempfile
from werkzeug.exceptions import HTTPException

from competition.registration.models import NormalizedCompetitor
from competition.registration.adapters.isoa_excel.reader import read_isoa_excel
from competition.registration.adapters.isoa_excel.mappers import map_row
from competition.registration.adapters.iof_xml.reader import read_iof_xml
from competition.registration.adapters.iof_xml.mappers import map_iof_competitor
from competition.registration.si_rental.validation import validate_si_rental_input
from competition.registration.si_rental.allocation import allocate_si_cards
from competition.registration.si_rental.excel import generate_si_rental_excel
from competition.registration.si_rental.api import competition


def _make_file_storage(xlsx_bytes, filename="startlist.xlsx"):
    class _FileStorage:
        def __init__(self):
            self.stream = io.BytesIO(xlsx_bytes)
            self.filename = filename

    return _FileStorage()


def _write_isoa_xlsx(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(list(row))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _call(starts, endings, missing, additional):
    return validate_si_rental_input(starts, endings, missing, additional)


# ---------------- validation (ported from ISOA test_competition.py) ----------------

def test_additional_cards_added_to_available_pool():
    req = _call(["100", "200"], ["105", "205"], ["", ""], "300, 301, 302")
    assert 100 in req.available_si_numbers
    assert 300 in req.available_si_numbers
    assert 301 in req.available_si_numbers
    assert 302 in req.available_si_numbers


def test_additional_cards_are_allocatable_when_no_boxes_filled():
    req = _call(["", ""], ["", ""], ["", ""], "500, 501")
    assert req.available_si_numbers == [500, 501]


def test_additional_cards_deduplicate_against_box_ranges():
    req = _call(["100"], ["105"], [""], "102, 103")
    assert req.available_si_numbers.count(102) == 1
    assert req.available_si_numbers.count(103) == 1


def test_additional_cards_whitespace_handled():
    req = _call(["100"], ["105"], [""], " 110 , 111 ")
    assert 110 in req.available_si_numbers
    assert 111 in req.available_si_numbers


def test_empty_additional_cards_string_not_added():
    req = _call(["100"], ["105"], [""], "")
    assert req.available_si_numbers == [100, 101, 102, 103, 104, 105]


def test_whitespace_only_additional_cards_not_added():
    req = _call(["100"], ["105"], [""], "   ")
    assert req.available_si_numbers == [100, 101, 102, 103, 104, 105]


def test_invalid_additional_card_aborts():
    with pytest.raises(HTTPException) as excinfo:
        _call(["100"], ["105"], [""], "abc")
    assert excinfo.value.code == 400


def test_additional_card_outside_range_still_added():
    req = _call(["100"], ["105"], [""], "999")
    assert 999 in req.available_si_numbers


def test_si_box_with_no_missing_cards_allowed():
    req = _call(["100"], ["105"], [""], "")
    assert req.available_si_numbers == [100, 101, 102, 103, 104, 105]


def test_additional_cards_returned_are_only_additional_not_box_ranges():
    req = _call(["100"], ["105"], [""], "200, 201")
    assert req.additional_cards == [200, 201]
    assert 100 not in req.additional_cards


def test_additional_cards_with_ranges_added_to_available_pool():
    req = _call(["100"], ["105"], [""], "200-202, 300")
    assert req.additional_cards == [200, 201, 202, 300]
    assert 200 in req.available_si_numbers
    assert 201 in req.available_si_numbers
    assert 202 in req.available_si_numbers
    assert 300 in req.available_si_numbers


def test_additional_cards_range_only_when_no_boxes_filled():
    req = _call(["", ""], ["", ""], ["", ""], "500-502")
    assert req.available_si_numbers == [500, 501, 502]


def test_additional_cards_range_deduplicate_against_box_ranges():
    req = _call(["100"], ["105"], [""], "102-104")
    assert req.available_si_numbers.count(102) == 1
    assert req.available_si_numbers.count(103) == 1
    assert req.available_si_numbers.count(104) == 1


def test_additional_cards_invalid_range_aborts():
    with pytest.raises(HTTPException) as excinfo:
        _call(["100"], ["105"], [""], "abc-def")
    assert excinfo.value.code == 400


def test_additional_cards_range_start_greater_than_end_aborts():
    with pytest.raises(HTTPException) as excinfo:
        _call(["100"], ["105"], [""], "200-190")
    assert excinfo.value.code == 400


def test_additional_cards_mixed_single_and_range():
    req = _call(["", ""], ["", ""], ["", ""], "1-3, 5, 8-10")
    assert req.available_si_numbers == [1, 2, 3, 5, 8, 9, 10]


# ---------------- validation: new branches ----------------

def test_empty_inputs_return_empty_pool():
    req = _call([], [], [], "")
    assert req.box_ranges == []
    assert req.missing_cards == []
    assert req.available_si_numbers == []
    assert req.additional_cards == []


def test_empty_box_row_skipped():
    req = _call(["", "100"], ["", "105"], ["", ""], "")
    assert req.available_si_numbers == [100, 101, 102, 103, 104, 105]


def test_box_start_not_digit_aborts():
    with pytest.raises(HTTPException) as excinfo:
        _call(["abc"], ["105"], [""], "")
    assert excinfo.value.code == 400


def test_box_end_not_digit_aborts():
    with pytest.raises(HTTPException) as excinfo:
        _call(["100"], ["xyz"], [""], "")
    assert excinfo.value.code == 400


def test_box_start_greater_than_end_aborts():
    with pytest.raises(HTTPException) as excinfo:
        _call(["200"], ["100"], [""], "")
    assert excinfo.value.code == 400


def test_missing_card_not_digit_aborts():
    with pytest.raises(HTTPException) as excinfo:
        _call(["100"], ["105"], ["abc"], "")
    assert excinfo.value.code == 400


def test_missing_card_outside_range_aborts():
    with pytest.raises(HTTPException) as excinfo:
        _call(["100"], ["105"], ["999"], "")
    assert excinfo.value.code == 400


def test_missing_cards_removed_from_available_pool():
    req = _call(["100"], ["105"], ["102, 104"], "")
    assert 102 not in req.available_si_numbers
    assert 104 not in req.available_si_numbers
    assert req.available_si_numbers == [100, 101, 103, 105]


def test_additional_token_not_digit_aborts():
    with pytest.raises(HTTPException) as excinfo:
        _call(["100"], ["105"], [""], "12x")
    assert excinfo.value.code == 400


def test_additional_empty_token_is_skipped():
    # empty token between commas hits the `continue` branch
    req = _call(["100"], ["105"], [""], "100,,101")
    assert 101 in req.available_si_numbers
    assert req.available_si_numbers.count(100) == 1


def test_additional_range_more_than_two_parts_aborts():
    with pytest.raises(HTTPException) as excinfo:
        _call(["100"], ["105"], [""], "1-2-3")
    assert excinfo.value.code == 400


# ---------------- allocation ----------------

def _competitor(member_id, name, rental=False, si_card=""):
    return NormalizedCompetitor(
        member_id=member_id, name=name, club="C", category="Cat",
        si_card_number=si_card, rental_requested=rental, note="", phone="",
    )


def test_allocate_assigns_to_rental_requested_in_order():
    comps = [_competitor("1", "A", True), _competitor("2", "B", True), _competitor("3", "C", False)]
    renders = allocate_si_cards(comps, [200, 201])
    assert [r.member_id for r in renders] == ["1", "2"]
    assert comps[0].si_card_number == "200"
    assert comps[1].si_card_number == "201"
    assert comps[2].si_card_number == ""


def test_allocate_skips_non_rental():
    comps = [_competitor("1", "A", False), _competitor("2", "B", True)]
    renders = allocate_si_cards(comps, [200])
    assert [r.member_id for r in renders] == ["2"]
    assert comps[0].si_card_number == ""
    assert comps[1].si_card_number == "200"


def test_allocate_stops_when_pool_exhausted():
    comps = [_competitor("1", "A", True), _competitor("2", "B", True), _competitor("3", "C", True)]
    renders = allocate_si_cards(comps, [200])
    assert [r.member_id for r in renders] == ["1"]
    assert comps[2].si_card_number == ""


def test_allocate_empty_pool_never_raises():
    comps = [_competitor("1", "A", True)]
    renders = allocate_si_cards(comps, [])
    assert renders == []
    assert comps[0].si_card_number == ""


def test_allocate_empty_competitors():
    assert allocate_si_cards([], [1, 2]) == []


# ---------------- adapter: isoa_excel mapper ----------------

def test_map_row_full():
    row = {
        "0": " 2071000 ", "1": " Cohen, Yosi ", "2": "Hapoel ", "3": "D21A ",
        "4": "", "5": "", "6": " 123456 ", "7": "", "8": "", "9": " note ",
        "10": "True ", "11": "", "12": "", "13": "", "14": "", "15": " 052-1234567 ",
    }
    c = map_row(row)
    assert c.member_id == "2071000"
    assert c.name == "Cohen, Yosi"
    assert c.club == "Hapoel"
    assert c.category == "D21A"
    assert c.si_card_number == "123456"
    assert c.rental_requested is True
    assert c.note == "note"
    assert c.phone == "052-1234567"


def test_map_row_rental_flag_one():
    row = {str(i): "" for i in range(16)}
    row["0"] = "1"
    row["1"] = "A"
    row["2"] = "C"
    row["3"] = "Cat"
    row["6"] = ""
    row["9"] = ""
    row["10"] = "1"
    row["15"] = ""
    c = map_row(row)
    assert c.rental_requested is True


def test_map_row_rental_flag_false():
    row = {str(i): "" for i in range(16)}
    row["10"] = "False"
    c = map_row(row)
    assert c.rental_requested is False


def test_map_row_rental_flag_empty():
    row = {str(i): "" for i in range(16)}
    row["10"] = ""
    c = map_row(row)
    assert c.rental_requested is False


# ---------------- adapter: isoa_excel reader ----------------

def test_reader_yields_dict_rows_keyed_by_index():
    row1 = [
        "2071000", "Cohen, Yosi", "Hapoel", "D21A", "", "", "123456", "", "", "note",
        "True", "", "", "", "", "052-123",
    ]
    row2 = [
        "2071001", "Levi, Sara", "Maccabi", "D20", "", "", "", "", "", "", "", "", "", "", "", "",
    ]
    xlsx = _write_isoa_xlsx([row1, row2])
    rows = list(read_isoa_excel(_make_file_storage(xlsx)))
    assert len(rows) == 2
    assert rows[0]["0"] == "2071000"
    assert rows[0]["1"] == "Cohen, Yosi"
    assert rows[0]["10"] == "True"
    assert rows[1]["15"] == ""


# ---------------- adapter: iof_xml ----------------

def _iof_xml_string():
    return b"""<?xml version="1.0"?>
<Result xmlns="http://www.orienteering.org/datastandard/3.0">
  <ClassResult>
    <Class><Name>D21A</Name></Class>
    <PersonResult>
      <Person><Id>100</Id><Name><Family>Cohen</Family><Given>Yosi</Given></Name></Person>
      <Result><Status>OK</Status><Time>3600</Time><BibNumber>5</BibNumber></Result>
    </PersonResult>
    <PersonResult>
      <Person><Id></Id><Name><Family>Unknown</Family><Given></Given></Name></Person>
      <Result><Status>OK</Status><Time>100</Time></Result>
    </PersonResult>
  </ClassResult>
</Result>"""


def test_iof_reader_yields_competitors():
    raws = list(read_iof_xml(_make_file_storage(_iof_xml_string(), "results.xml")))
    assert len(raws) == 2
    assert raws[0]["class_name"] == "D21A"
    assert raws[0]["element"] is not None


def test_iof_mapper_full():
    raws = list(read_iof_xml(_make_file_storage(_iof_xml_string(), "results.xml")))
    c = map_iof_competitor(raws[0])
    assert c.member_id == "100"
    assert c.name == "Yosi Cohen"
    assert c.category == "D21A"
    assert c.rental_requested is False


def test_iof_mapper_falls_back_to_bib_and_empty_id():
    raws = list(read_iof_xml(_make_file_storage(_iof_xml_string(), "results.xml")))
    c = map_iof_competitor(raws[1])
    assert c.member_id == ""
    assert c.name == "Unknown"
    assert c.category == "D21A"


# ---------------- excel ----------------

def test_generate_si_rental_excel_returns_xlsx_bytes():
    box_ranges = [range(100, 103)]
    missing = [101]
    renders = [
        NormalizedCompetitor("1", "A", "C", "Cat", "100", True, "", "052-1"),
        NormalizedCompetitor("2", "B", "C", "Cat", "102", True, "", "052-2"),
    ]
    data = generate_si_rental_excel(box_ranges, missing, renders, [])
    assert data.startswith(b"PK")  # xlsx zip signature
    wb = openpyxl.load_workbook(io.BytesIO(data))
    assert "מזוודה 1" in wb.sheetnames
    ws = wb["מזוודה 1"]
    assert ws["A5"].value == "100"
    assert ws["B5"].value == "1"


def test_generate_si_rental_excel_additional_sheet():
    data = generate_si_rental_excel([range(100, 101)], [], [], [200, 201])
    wb = openpyxl.load_workbook(io.BytesIO(data))
    assert "מזוודה 1" in wb.sheetnames
    assert "כרטיסים נוספים" in wb.sheetnames


def test_generate_si_rental_excel_missing_card_red_rows():
    data = generate_si_rental_excel([range(100, 102)], [100], [], [])
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb["מזוודה 1"]
    assert ws["A5"].value == "100"


# ---------------- api ----------------

def _make_app():
    from flask import Flask
    app = Flask(__name__)
    app.register_blueprint(competition, url_prefix="/api/competition")
    return app


def _xlsx_file():
    return _make_file_storage(_write_isoa_xlsx([
        ["2071000", "Cohen, Yosi", "Hapoel", "D21A", "", "", "", "", "", "", "True", "", "", "", "", "052-1"],
    ]))


def test_si_rental_endpoint_returns_zip():
    app = _make_app()
    client = app.test_client()
    data = {
        "file": (io.BytesIO(_write_isoa_xlsx([
            ["2071000", "Cohen, Yosi", "Hapoel", "D21A", "", "", "", "", "", "", "True", "", "", "", "", "052-1"],
        ])), "startlist.xlsx"),
        "boxStart[]": ["100"],
        "boxEnd[]": ["105"],
        "missingCards[]": [""],
        "additionalSiCards": "",
    }
    resp = client.post("/api/competition/si-rental", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200
    assert resp.mimetype == "application/zip"
    assert resp.content_type.startswith("application/zip")
    body = resp.get_data()
    assert body.startswith(b"PK")


def test_si_rental_endpoint_missing_file_returns_json_400():
    app = _make_app()
    client = app.test_client()
    resp = client.post("/api/competition/si-rental", data={}, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert resp.is_json
    assert resp.get_json() == {"error": "Startlist file is required."}


def test_si_rental_endpoint_invalid_box_range_returns_json_400():
    app = _make_app()
    client = app.test_client()
    data = {
        "file": (io.BytesIO(_write_isoa_xlsx([
            ["2071000", "Cohen, Yosi", "Hapoel", "D21A", "", "", "", "", "", "", "True", "", "", "", "", "052-1"],
        ])), "startlist.xlsx"),
        "boxStart[]": ["abc"],
        "boxEnd[]": ["105"],
        "missingCards[]": [""],
        "additionalSiCards": "",
    }
    resp = client.post("/api/competition/si-rental", data=data, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert resp.is_json
    assert "error" in resp.get_json()


def test_si_rental_endpoint_invalid_extension_returns_json_400():
    app = _make_app()
    client = app.test_client()
    data = {
        "file": (io.BytesIO(b"not an xlsx"), "startlist.txt"),
        "boxStart[]": ["100"],
        "boxEnd[]": ["105"],
        "missingCards[]": [""],
        "additionalSiCards": "",
    }
    resp = client.post("/api/competition/si-rental", data=data, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert resp.is_json
    assert "error" in resp.get_json()


def test_si_rental_endpoint_empty_filename_returns_json_400():
    app = _make_app()
    client = app.test_client()
    data = {
        "file": (io.BytesIO(b""), ""),
        "boxStart[]": ["100"],
        "boxEnd[]": ["105"],
        "missingCards[]": [""],
        "additionalSiCards": "",
    }
    resp = client.post("/api/competition/si-rental", data=data, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert resp.is_json
    assert resp.get_json() == {"error": "Startlist file is required."}


# ---------------- write_start_file: SI rental card + blank lines ----------------


def _write_start_file_csv(competitors, si_card_map=None):
    d = tempfile.mkdtemp()
    path = app_module.write_start_file(competitors, d + "/", si_card_map=si_card_map)
    with open(path, "rb") as f:
        return f.read()


def test_write_start_file_uses_allocated_si_card_for_rental():
    competitors = [
        ["Shorty", 100, "Member With Card", "C", "D21A", None, None, 500000, "052-1"],
        ["Shorty", 20001, "Rental No Card", "C", "D21A", None, None, None, "052-2"],
        ["Shorty", 20002, "Rental Low Card", "C", "D21A", None, None, 999, "052-3"],
    ]
    si_card_map = {"20001": "300", "20002": "301"}
    raw = _write_start_file_csv(competitors, si_card_map)
    text = raw.decode("cp1255")
    rows = list(csv.reader(io.StringIO(text)))
    assert rows[0][5] == "CARD NUMBER"
    # member keeps own card
    assert rows[1][5] == "500000"
    # rental with no card gets allocated card
    assert rows[2][5] == "300"
    # rental with low card gets allocated card
    assert rows[3][5] == "301"
    # rental tag present for rentals, absent for member
    assert rows[2][7] == "rental"
    assert rows[3][7] == "rental"
    assert rows[1][7] == ""


def test_write_start_file_has_no_blank_lines():
    competitors = [
        ["Shorty", 100, "Member With Card", "C", "D21A", None, None, 500000, "052-1"],
        ["Shorty", 20001, "Rental No Card", "C", "D21A", None, None, None, "052-2"],
    ]
    raw = _write_start_file_csv(competitors, None)
    # no double line endings that render as blank lines
    assert b"\r\r\n" not in raw
    assert b"\n\n" not in raw
    text = raw.decode("cp1255")
    assert "" not in text.splitlines()


def test_write_start_file_backwards_compatible_without_si_card_map():
    competitors = [
        ["Shorty", 100, "Member With Card", "C", "D21A", None, None, 500000, "052-1"],
    ]
    raw = _write_start_file_csv(competitors, None)
    text = raw.decode("cp1255")
    rows = list(csv.reader(io.StringIO(text)))
    assert rows[0][5] == "CARD NUMBER"
    assert rows[1][5] == "500000"
    assert rows[1][7] == ""


def _apply_si_card_map(competitors, si_card_map):
    """Mirror Start_Draw.py: propagate allocated SI cards to comp_list card numbers."""
    for runner in competitors:
        if str(runner[1]) in si_card_map and (runner[7] is None or runner[7] < 100000):
            runner[7] = int(si_card_map[str(runner[1])])
            runner.append(True)  # mark allocated rental SI for HTML marking


def _write_html(competitors, si_card_map=None, fn_name="write_html_file_by_category"):
    d = tempfile.mkdtemp()
    _apply_si_card_map(competitors, si_card_map or {})
    getattr(app_module, fn_name)(competitors, d + "/")
    path = d + "/" + (
        "HTML_Start_Times_By_Category.html"
        if fn_name == "write_html_file_by_category"
        else "HTML_Start_Times_By_Starting_Time.html"
    )
    with open(path, "r", encoding="utf8") as f:
        return f.read()


def test_html_by_category_shows_allocated_si_for_rental():
    competitors = [
        ["Shorty", 100, "Member With Card", "C", "D21A", None, None, 500000, "052-1"],
        ["Shorty", 20001, "Rental No Card", "C", "D21A", None, None, None, "052-2"],
        ["Shorty", 20002, "Rental Low Card", "C", "D21A", None, None, 999, "052-3"],
    ]
    si_card_map = {"20001": "300", "20002": "301"}
    html = _write_html(competitors, si_card_map, "write_html_file_by_category")
    # member keeps own card; rentals show allocated cards
    assert "500000" in html
    assert "300" in html
    assert "301" in html


def test_html_by_starting_time_shows_allocated_si_for_rental():
    competitors = [
        ["Shorty", 100, "Member With Card", "C", "D21A", datetime.time(9, 0), None, 500000, "052-1"],
        ["Shorty", 20001, "Rental No Card", "C", "D21A", datetime.time(9, 5), None, None, "052-2"],
    ]
    si_card_map = {"20001": "300"}
    html = _write_html(competitors, si_card_map, "write_html_file_by_starting_time")
    assert "500000" in html
    assert "300" in html


def test_html_by_category_marks_rental_si():
    competitors = [
        ["Shorty", 100, "Member With Card", "C", "D21A", None, None, 500000, "052-1"],
        ["Shorty", 20001, "Rental No Card", "C", "D21A", None, None, None, "052-2"],
    ]
    si_card_map = {"20001": "300"}
    html = _write_html(competitors, si_card_map, "write_html_file_by_category")
    # member card not wrapped in rental marker
    assert '<span class="si-rental" title="כרטיס שכור">500000</span>' not in html
    # rental card wrapped in rental marker
    assert '<span class="si-rental" title="כרטיס שכור">300</span>' in html
    assert 'class="si-rental"' in html


def test_html_by_starting_time_marks_rental_si():
    competitors = [
        ["Shorty", 100, "Member With Card", "C", "D21A", datetime.time(9, 0), None, 500000, "052-1"],
        ["Shorty", 20001, "Rental No Card", "C", "D21A", datetime.time(9, 5), None, None, "052-2"],
    ]
    si_card_map = {"20001": "300"}
    html = _write_html(competitors, si_card_map, "write_html_file_by_starting_time")
    assert '<span class="si-rental" title="כרטיס שכור">500000</span>' not in html
    assert '<span class="si-rental" title="כרטיס שכור">300</span>' in html
    assert 'class="si-rental"' in html


import pytest

# ---------------- end-to-end: real startlist xlsx -> StartList.csv contains rental ----------------


def _competitors_from_xlsx(path):
    """Read an ISOA-produced StartList xlsx and build competitors for write_start_file.

    Columns: A=STNO, B=NAME, C=CLUB, D=CLASS NAME, E=requested time,
    G=CARD NUMBER, P=PHONE.
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    competitors = []
    for row in range(2, ws.max_row + 1):
        stno = ws.cell(row=row, column=1).value
        if stno is None:
            continue
        competitors.append([
            "Test",                      # course
            int(stno),                   # STNO (int, compared to min_external_id)
            ws.cell(row=row, column=2).value,  # NAME
            ws.cell(row=row, column=3).value,  # CLUB
            ws.cell(row=row, column=4).value,  # CLASS NAME
            ws.cell(row=row, column=5).value,  # requested start time
            None,                        # start time (filled by draw)
            ws.cell(row=row, column=7).value,  # CARD NUMBER
            ws.cell(row=row, column=16).value, # PHONE
        ])
    return competitors


RENTAL_FIXTURES = [
    "tests/fixtures/isoa_startlist_with_rental.xlsx",
    "tests/fixtures/isoa_startlist_rental_input_19.xlsx",
    "tests/fixtures/isoa_startlist_rental_input_20.xlsx",
    "tests/fixtures/isoa_startlist_rental_input_21.xlsx",
]


@pytest.mark.parametrize("path", RENTAL_FIXTURES)
def test_startlist_csv_contains_rental_from_real_file(path):
    import copy
    competitors = _competitors_from_xlsx(path)
    # at least one row has a missing/low card number -> should become rental
    rental_candidates = [c for c in competitors if c[7] is None or c[7] < 100000]
    assert rental_candidates, "fixture %s should contain a rental candidate" % path

    # without si_card_map: card None/<100000 still yields rental memo
    raw = _write_start_file_csv(copy.deepcopy(competitors), None)
    text = raw.decode("cp1255")
    assert "rental" in text, "startlist.csv should mark rented SI cards with rental memo for %s" % path

    # with si_card_map: allocated card is written and rental memo is set
    si_card_map = {str(c[1]): "300" + str(c[1])[-3:] for c in rental_candidates}
    raw2 = _write_start_file_csv(copy.deepcopy(competitors), si_card_map)
    text2 = raw2.decode("cp1255")
    assert "rental" in text2, "startlist.csv should mark rented SI cards with rental memo for %s" % path
    for c in rental_candidates:
        allocated = "300" + str(c[1])[-3:]
        assert allocated in text2, "allocated SI card %s should appear in CSV for %s" % (allocated, path)

