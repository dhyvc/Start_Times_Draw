from flask import Flask, flash, redirect, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import app as app_module
import datetime
import os
import io

from competition.registration.adapters.isoa_excel.reader import read_isoa_excel
from competition.registration.adapters.isoa_excel.mappers import map_row
from competition.registration.adapters.isoa_excel.validation import validate_rows
from competition.registration.adapters.isoa_excel.xlsx import generate_validation_xlsx
from competition.registration.si_rental.validation import validate_si_rental_input
from competition.registration.si_rental.allocation import allocate_si_cards
from competition.registration.si_rental.excel import generate_si_rental_excel

app = Flask(__name__)
app.secret_key = '7ff1db0ee6826e68c0ae9abdd8da1966fb7cbdbf5dbe7c77915d4cd10b9e7999'
app.config["IMAGE_UPLOADS"] = "./uploads"
app.config["IMAGE_DOWNLOADS"] = "./downloads"

@app.route('/', methods = ['POST', 'GET'])
def index():
	if request.method == 'POST':
# check if the post request has the file part
		if 'event_file' not in request.files:
			flash('No file part')
			return redirect(request.url)

		event_type = request.form.get("event_type")
		if event_type is None or event_type == "":
			flash('Missing Event Type!')
			return redirect(request.url)

		iterations = request.form.get("iterations")
		if iterations is None or iterations == "":
			flash('Missing Iterations!')
			return redirect(request.url)

		same_club_weight = request.form.get("same_club_weight")
		if same_club_weight is None or same_club_weight == "":
			flash('Missing Club Weight!')
			return redirect(request.url)

		same_category_weight = request.form.get("same_category_weight")
		if same_category_weight is None or same_category_weight == "":
			flash('Missing Category Weight!')
			return redirect(request.url)

		first_start = request.form.get("first_start")
		if first_start is None or first_start == "":
			flash('Missing First Start Time!')
			return redirect(request.url)

		last_start = request.form.get("last_start")
		if last_start is None or last_start == "":
			flash('Missing Last Start Time!')
			return redirect(request.url)

		start_window = request.form.get("start_window")
		if start_window is None or start_window == "":
			flash('Missing Start Window Size!')
			return redirect(request.url)

		vacant_slot = request.form.get("vacant_slot")
		if vacant_slot is None or vacant_slot == "":
			flash('Missing Vacant Slot Interval!')
			return redirect(request.url)

		event_file = request.files["event_file"]
		et = event_type
		print ("event type: " + event_type)
		dt_format = "%H:%M"
		fs = (datetime.datetime.strptime(first_start, dt_format)).time()
		ls = datetime.datetime.strptime(last_start, dt_format).time()
		sw = int(start_window)
		vs = int(vacant_slot)
		oi = int(iterations)
		ccl = int(same_club_weight)
		cca = int(same_category_weight)
		# print(sw)
		# sw = datetime.datetime.strptime(start_window, dt_format).time()

		# sane = app.sanity_check(fs, ls, sw.minute)[3]

		fs, ls, sw, err = app_module.sanity_check(fs, ls, sw)
		# fs, ls, sw, err = app.sanity_check(fs, ls, sw.minute)
		# print(secure_filename(event_file.filename))

		os.makedirs(app.config["IMAGE_UPLOADS"], exist_ok=True)
		os.makedirs(app.config["IMAGE_DOWNLOADS"], exist_ok=True)

		event_file.save(os.path.join(app.config["IMAGE_UPLOADS"], secure_filename(event_file.filename)))

		event_file_path = os.path.join(app.config["IMAGE_UPLOADS"], secure_filename(event_file.filename))

		comp_list = app_module.read_start_file(event_file_path, fs, ls, sw, vs, et, oi, ccl, cca)

		# SI Card Rental: generate SI rental sheets as part of the draw flow when SI box data is provided,
		# reusing the event Excel file already uploaded.
		si_card_map = None
		if request.form.get("si_rental") or request.form.getlist("boxStart[]") or request.form.getlist("boxEnd[]") or request.form.getlist("missingCards[]") or request.form.get("additionalSiCards", "").strip() != "":
			box_starts = request.form.getlist("boxStart[]")
			box_ends = request.form.getlist("boxEnd[]")
			missing = request.form.getlist("missingCards[]")
			additional = request.form.get("additionalSiCards", "")

			if box_starts or box_ends or missing or additional.strip() != "":
				rental_request = validate_si_rental_input(box_starts, box_ends, missing, additional)
				with open(event_file_path, "rb") as f:
					si_storage = type("FS", (), {"stream": io.BytesIO(f.read())})()
				si_competitors = [map_row(row) for row in read_isoa_excel(si_storage)]
				renders = allocate_si_cards(si_competitors, rental_request.available_si_numbers)
				si_card_map = {c.member_id: c.si_card_number for c in renders}
				# Propagate allocated SI cards to the start list so they appear
				# in the CSV and all HTML startlist outputs for rental runners.
				for runner in comp_list:
					if str(runner[1]) in si_card_map and (runner[7] is None or runner[7] < 100000):
						runner[7] = int(si_card_map[str(runner[1])])
						runner.append(True)  # mark allocated rental SI for HTML marking
				xlsx_bytes = generate_si_rental_excel(
					rental_request.box_ranges,
					rental_request.missing_cards,
					renders,
					rental_request.additional_cards,
				)
				with open("./downloads/SI_Rental.xlsx", "wb") as f:
					f.write(xlsx_bytes)

		app_module.write_start_file(comp_list, "./downloads/", si_card_map=si_card_map)
		app_module.write_undefined_registrations(comp_list, "./downloads/")
		app_module.write_html_file_for_website(comp_list, "./downloads/")
		app_module.write_html_file_by_category(comp_list, "./downloads/")
		app_module.write_html_file_by_starting_time(comp_list, "./downloads/")
		app_module.write_vacant_slots_by_course(comp_list, "./downloads/", fs, ls, et)

		if os.path.isfile("./downloads/StartList.zip"):
			os.remove("./downloads/StartList.zip")
			dl_list = [f for f in os.listdir("./downloads") if f != "ISOA_Validation.xlsx" and not f.lower().endswith(".pdf")]
			app_module.make_zip_file('./downloads', dl_list)
			# return redirect(request.url)
		else:
			print ('does not exist')
			dl_list = [f for f in os.listdir("./downloads") if f != "ISOA_Validation.xlsx" and not f.lower().endswith(".pdf")]
			app_module.make_zip_file('./downloads', dl_list)
		print(dl_list)
		return send_file('downloads/StartList.zip', download_name="StartList.zip")
		# return send_file("./uploads/StartList.xlsx")
		# 
	else:
		return render_template('start_draw.html')

@app.route('/validate-isoa', methods=['POST'])
def validate_isoa():
	"""Validate an uploaded ISOA Excel startlist and return an XLSX report
	describing every problem found with reasoning."""
	if 'isoa_file' not in request.files:
		return jsonify({"error": "No file part"}), 400

	isoa_file = request.files['isoa_file']
	if isoa_file.filename == '':
		return jsonify({"error": "No selected file"}), 400

	filename = secure_filename(isoa_file.filename)
	if not filename.lower().endswith('.xlsx'):
		return jsonify({"error": "Invalid file. Only .xlsx files are allowed."}), 400

	rows = list(read_isoa_excel(isoa_file))
	problems = validate_rows(rows)
	xlsx_bytes = generate_validation_xlsx(problems)

	out_path = os.path.join(app.config["IMAGE_DOWNLOADS"], "ISOA_Validation.xlsx")
	if not os.path.exists(app.config["IMAGE_DOWNLOADS"]):
		os.makedirs(app.config["IMAGE_DOWNLOADS"])
	with open(out_path, "wb") as f:
		f.write(xlsx_bytes)

	xlsx_mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
	return send_file(out_path, mimetype=xlsx_mimetype,
					 as_attachment=True, download_name="ISOA_Validation.xlsx")


@app.route('/about')
def about():
	return render_template('about.html')
