from flask import Flask, flash, redirect, render_template, request, send_file
from werkzeug.utils import secure_filename
import my_main
import datetime
import os

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
		# print(sw)
		# sw = datetime.datetime.strptime(start_window, dt_format).time()

		# sane = my_main.sanity_check(fs, ls, sw.minute)[3]

		fs, ls, sw, err = my_main.sanity_check(fs, ls, sw)
		# fs, ls, sw, err = my_main.sanity_check(fs, ls, sw.minute)
		# print(secure_filename(event_file.filename))

		event_file.save(os.path.join(app.config["IMAGE_UPLOADS"], secure_filename(event_file.filename)))

		event_file_path = os.path.join(app.config["IMAGE_UPLOADS"], event_file.filename)

		comp_list = my_main.read_start_file(event_file_path, fs, ls, sw, vs)

		my_main.write_start_file(comp_list, "./downloads/")
		my_main.write_html_file_by_category(comp_list, "./downloads/")
		my_main.write_html_file_by_starting_time(comp_list, "./downloads/")
		my_main.write_vacant_slots_by_course(comp_list, "./downloads/", fs, ls)

		if os.path.isfile("./downloads/StartList.zip"):
			os.remove("./downloads/StartList.zip")
			dl_list = os.listdir("./downloads")
			my_main.make_zip_file('./downloads', dl_list)
			# return redirect(request.url)
		else:
			print ('does not exist')
			dl_list = os.listdir("./downloads")
			my_main.make_zip_file('./downloads', dl_list)
		print(dl_list)
		return send_file('downloads/StartList.zip', download_name="StartList.zip")
		# return send_file("./uploads/StartList.xlsx")
		# 
	else:
		return render_template('start_draw.html')

@app.route('/about')
def about():
	return render_template('about.html')
