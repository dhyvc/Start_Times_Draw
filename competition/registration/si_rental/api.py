import os
import io
from zipfile import ZipFile

from flask import Blueprint, jsonify, request, send_file, abort
from werkzeug.utils import secure_filename

from competition.registration.adapters.isoa_excel.reader import read_isoa_excel
from competition.registration.adapters.isoa_excel.mappers import map_row
from competition.registration.si_rental.validation import validate_si_rental_input
from competition.registration.si_rental.allocation import allocate_si_cards
from competition.registration.si_rental.excel import generate_si_rental_excel

competition = Blueprint('competition', __name__)

ALLOWED_FILE_EXTENSIONS = [".xlsx"]


@competition.app_errorhandler(400)
def handle_bad_request(e):
    return jsonify({"error": e.description or "Bad Request"}), 400


def _validate_startlist_file(file_storage):
    filename = secure_filename(file_storage.filename) if file_storage and file_storage.filename else ""
    if filename == "":
        abort(400, "Startlist file is required.")
    file_ext = os.path.splitext(filename)[1]
    if file_ext not in ALLOWED_FILE_EXTENSIONS:
        abort(400, "Invalid startlist file. Only .xlsx files are allowed.")
    return file_storage


@competition.route("/si-rental", methods=["POST"])
def si_rental():
    """POST /api/competition/si-rental (multipart form-data).

    Form fields: file, boxStart[], boxEnd[], missingCards[], additionalSiCards.
    Returns a zip file download containing the SI rental xlsx, or a JSON error
    with status 400 on invalid input.
    """
    if "file" not in request.files:
        abort(400, "Startlist file is required.")
    file_storage = _validate_startlist_file(request.files["file"])

    box_starts = request.form.getlist("boxStart[]")
    box_ends = request.form.getlist("boxEnd[]")
    missing = request.form.getlist("missingCards[]")
    additional = request.form.get("additionalSiCards", "")

    rental_request = validate_si_rental_input(box_starts, box_ends, missing, additional)

    competitors = [map_row(row) for row in read_isoa_excel(file_storage)]
    renders = allocate_si_cards(competitors, rental_request.available_si_numbers)

    xlsx_bytes = generate_si_rental_excel(
        rental_request.box_ranges,
        rental_request.missing_cards,
        renders,
        rental_request.additional_cards,
    )

    zip_buffer = io.BytesIO()
    with ZipFile(zip_buffer, "w") as zf:
        zf.writestr("השאלת_כרטיסי_SI.xlsx", xlsx_bytes)
    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name="si_rental.zip",
    )
