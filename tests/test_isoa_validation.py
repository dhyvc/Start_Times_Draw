import io

import openpyxl

import Start_Draw
from competition.registration.adapters.isoa_excel.reader import read_isoa_excel
from competition.registration.adapters.isoa_excel.validation import validate_rows
from competition.registration.adapters.isoa_excel.xlsx import generate_validation_xlsx


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


def _rows(xlsx_bytes):
    return list(read_isoa_excel(_make_file_storage(xlsx_bytes)))


def _row(member_id="2071000", name="Cohen, Yosi", club="Hapoel", category="D21A",
         si_card="123456", note="", rental="", phone="052-1"):
    # Column index layout: 0=member,1=name,2=club,3=class,6=SI card,9=note,10=rental,15=phone
    cols = [""] * 16
    cols[0] = member_id
    cols[1] = name
    cols[2] = club
    cols[3] = category
    cols[6] = si_card
    cols[9] = note
    cols[10] = rental
    cols[15] = phone
    return cols


# ---------------- validation rules ----------------

def test_undefined_category_reported():
    xlsx = _write_isoa_xlsx([_row(category="")])
    problems = validate_rows(_rows(xlsx))
    assert len(problems) == 1
    assert problems[0].member_id == "2071000"
    assert problems[0].category == ""
    assert "Undefined registration" in problems[0].reason


def test_valid_comment_not_flagged():
    xlsx = _write_isoa_xlsx([_row(note="[מנוי ליגה]")])
    problems = validate_rows(_rows(xlsx))
    assert problems == []


def test_empty_comment_not_flagged():
    xlsx = _write_isoa_xlsx([_row(note="")])
    problems = validate_rows(_rows(xlsx))
    assert problems == []


def test_non_empty_comment_reported():
    xlsx = _write_isoa_xlsx([_row(note="Please call me")])
    problems = validate_rows(_rows(xlsx))
    assert len(problems) == 1
    assert problems[0].reason == "Non standard comment"
    assert problems[0].note == "Please call me"


def test_card_zero_without_rental_reported():
    xlsx = _write_isoa_xlsx([_row(si_card="0", rental="")])
    problems = validate_rows(_rows(xlsx))
    assert len(problems) == 1
    assert "card number is 0" in problems[0].reason


def test_card_zero_with_rental_not_flagged():
    xlsx = _write_isoa_xlsx([_row(si_card="0", rental="True")])
    problems = validate_rows(_rows(xlsx))
    assert problems == []


def test_nonzero_card_not_flagged():
    xlsx = _write_isoa_xlsx([_row(si_card="123456", rental="")])
    problems = validate_rows(_rows(xlsx))
    assert problems == []


def test_clean_registration_no_problems():
    xlsx = _write_isoa_xlsx([_row()])
    problems = validate_rows(_rows(xlsx))
    assert problems == []


def test_multiple_rules_on_same_row():
    xlsx = _write_isoa_xlsx([_row(category="", note="hello", si_card="0", rental="")])
    problems = validate_rows(_rows(xlsx))
    reasons = {p.reason for p in problems}
    assert len(problems) == 3
    assert any("Undefined registration" in r for r in reasons)
    assert any("Non standard comment" in r for r in reasons)
    assert any("card number is 0" in r for r in reasons)


# ---------------- XLSX ----------------

def test_generate_validation_xlsx_returns_xlsx_bytes():
    problems = validate_rows(_rows(_write_isoa_xlsx([_row(note="hello")])))
    data = generate_validation_xlsx(problems)
    assert data.startswith(b"PK")  # xlsx zip signature
    assert len(data) > 1000


def test_generate_validation_xlsx_no_problems():
    data = generate_validation_xlsx([])
    assert data.startswith(b"PK")


# ---------------- route: /validate-isoa ----------------

def _make_app_client():
    return Start_Draw.app.test_client()


def test_validate_isoa_route_returns_xlsx():
    xlsx = _write_isoa_xlsx([_row(note="hello")])
    client = _make_app_client()
    data = {
        "isoa_file": (io.BytesIO(xlsx), "startlist.xlsx"),
    }
    resp = client.post("/validate-isoa", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200
    assert "spreadsheetml" in resp.mimetype
    body = resp.get_data()
    assert body.startswith(b"PK")


def test_validate_isoa_route_missing_file_returns_400():
    client = _make_app_client()
    resp = client.post("/validate-isoa", data={}, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert resp.is_json
    assert "error" in resp.get_json()


def test_validate_isoa_route_invalid_extension_returns_400():
    client = _make_app_client()
    data = {
        "isoa_file": (io.BytesIO(b"not an xlsx"), "startlist.txt"),
    }
    resp = client.post("/validate-isoa", data=data, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert resp.is_json
    assert "error" in resp.get_json()
