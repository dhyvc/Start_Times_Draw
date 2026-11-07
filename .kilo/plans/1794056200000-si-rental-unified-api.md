# Plan: Unified SI Rental API for Start_Times_Draw

## Goal
Add SI rental as a utility behind a shared competition API (Flask Blueprint), operating on a normalized competitor contract independent of input source. Keep the existing start-times-draw UI and logic unchanged — add a single "SI Rental" checkbox to the current form that reveals the SI rental fields (box icon kept). No tabs.

## Decisions (resolved)
- Scope = SI rental only; start-times-draw left untouched.
- Package structure under `competition/registration/`:
  - `models.py` (shared contract)
  - `adapters/` — input adapters; each source has `reader.py` + `mappers.py` (+ optional `generated/`)
    - `adapters/isoa_excel/reader.py`, `adapters/isoa_excel/mappers.py` — ISOA startlist.xlsx
    - `adapters/iof_xml/reader.py`, `adapters/iof_xml/mappers.py`, `adapters/iof_xml/generated/` — IOF results xml
  - `si_rental/` — SI rental domain (validation, allocation, excel, api)
- Blueprint `competition` defined in `competition/registration/si_rental/api.py`, registered at url_prefix `/api/competition`.
- SI rental endpoint: `POST /api/competition/si-rental` (multipart form-data), returns zip file download or JSON error.
- UI = existing `templates/start_draw.html` with one "SI Rental" checkbox; checked -> show SI box/missing/additional fields; box icon retained. No tabs, no separate page.
- Dependencies added: `pandas`, `xlsxwriter`.

## Files to create / modify

### New: `competition/__init__.py`
- Package marker.

### New: `competition/registration/__init__.py`
- Package marker.

### New: `competition/registration/models.py`
- `NormalizedCompetitor` dataclass (shared contract): `member_id, name, club, category, si_card_number, rental_requested, note, phone`.

### New: `competition/registration/adapters/__init__.py`
- Package marker.

### New: `competition/registration/adapters/isoa_excel/__init__.py`
- Package marker.

### New: `competition/registration/adapters/isoa_excel/reader.py`
- Opens ISOA startlist.xlsx workbook, yields raw dict rows.

### New: `competition/registration/adapters/isoa_excel/mappers.py`
- Maps raw excel row -> `NormalizedCompetitor`.

### New: `competition/registration/adapters/iof_xml/__init__.py`
- Package marker.

### New: `competition/registration/adapters/iof_xml/reader.py`
- Parses IOF results xml, yields raw elements / generated objects.

### New: `competition/registration/adapters/iof_xml/mappers.py`
- Maps IOF generated object -> `NormalizedCompetitor`.

### New: `competition/registration/adapters/iof_xml/generated/`
- Classes generated from IOF XSD (optional, if using xsdata).

### New: `competition/registration/si_rental/__init__.py`
- Package marker.

### New: `competition/registration/si_rental/validation.py`
- `validate_si_rental_input(box_starts, box_ends, missing, additional)` -> `SIRentalRequest` (abort 400 on invalid; port from ISOA `competition.py`).

### New: `competition/registration/si_rental/allocation.py`
- `allocate_si_cards(competitors, available_si_numbers)` -> assigns SI numbers to rental-requested competitors.

### New: `competition/registration/si_rental/excel.py`
- `generate_si_rental_excel(...)` -> SI rental xlsx bytes.

### New: `competition/registration/si_rental/api.py`
- Flask `Blueprint('competition', __name__)`.
- `POST /api/competition/si-rental`: validate SI input, translate file -> competitors, allocate SI cards, build excel, zip, return `send_file` attachment; errors -> JSON `{error}` 400.
- Docstring request/response contract.

### Modify: `templates/start_draw.html`
- Add "SI Rental" checkbox to existing form; checked shows SI fields (box start/end, missing cards, additional cards); keep suitcase icon.

### Modify: `Start_Draw.py`
- Register `competition` blueprint at `/api/competition`.
- Keep existing `/` route unchanged.

### Modify: `pyproject.toml`
- Add `pandas`, `xlsxwriter` to `dependencies`.
- Add `pytest`, `pytest-cov` to `dev` dependencies.

### New: `tests/test_si_rental.py`
- Port ISOA `test_competition.py` cases to normalized contract (`validate_si_rental_input`, `allocate_si_cards`).

## Data flow (SI rental)
1. Client submits form (SI Rental checkbox checked) -> POST `/api/competition/si-rental` multipart: `file`, `boxStart[]`, `boxEnd[]`, `missingCards[]`, `additionalSiCards`.
2. Blueprint validates SI input (abort 400 -> JSON error).
3. Chosen adapter converts uploaded file -> `List[NormalizedCompetitor]`.
4. `allocate_si_cards` assigns SI numbers to `rental_requested` competitors.
5. `generate_si_rental_excel` builds xlsx; results zipped; zip streamed as attachment.

## Failure modes
- Missing/invalid file -> 400 JSON.
- Non-numeric box ranges / start > end / missing card outside range -> 400 JSON.
- No available SI numbers -> allocation stops (existing behavior).
- Unimplemented adapter -> 400/501 documented.

## Validation
- `pytest tests/test_si_rental.py` passes with **100% test coverage** (measured via `pytest --cov=competition --cov-report=term-missing`; coverage threshold 100%).
- Flask test-client smoke test: POST `/api/competition/si-rental` returns 200 zip or 400 JSON.
- `python -m py_compile` on new files.