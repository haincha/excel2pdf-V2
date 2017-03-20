import sys
import os
import random
import time
from flask import Flask, request, render_template, jsonify, make_response, send_file, flash, url_for, redirect, Markup, Response
import pyexcel
import HTML
import pdfkit
import zipfile
import datetime
from celery import Celery
import pickle

app = Flask(__name__)
app.secret_key = 'some_secret'
app.config['CELERY_BROKER_URL'] = 'redis://localhost:6379/0'
app.config['CELERY_RESULT_BACKEND'] = 'redis://localhost:6379/0'
app.config['UPLOAD_FOLDER'] = '.'

celery = Celery(app.name, broker=app.config['CELERY_BROKER_URL'])
celery.conf.update(app.config)

@celery.task(bind=True)
def long_task(self):
	try:
		filename = ""
		waitingforfile = True
		while waitingforfile == True:
			filename = sorted(os.listdir(app.config['UPLOAD_FOLDER']), key=os.path.getctime)[-1:][0]
			if filename != "" and (filename.split(".")[-1:][0] == "xlsx" or filename.split(".")[-1:][0] == "xls"):
				waitingforfile = False
		time.sleep(3)
		with open ("{}.pckl".format(filename.split(".")[0]), 'rb') as fp:
			filename, accountlist, starttab, colname = pickle.load(fp)
		wb = pyexcel.get_book(file_name=str(os.path.join(app.config['UPLOAD_FOLDER'], filename)))
		sheets = wb.to_dict()
		all_sheets = []
		styling = 'display: inline; page-break-before: auto; padding-bottom: 50%; font-family: Calibri; font-size: 8.76;'
		found_accounts = 0
		account_column = 0
		header_column = 0
		header_row = 0
		header_found = False
		headerrow = []
		accountrow = []
		for name in sheets.keys():
			all_sheets.append(name)
		if isinstance(starttab, int) == False:
			starttab = 0
		if starttab > len(all_sheets):
			starttab = 0
		if starttab > 1:
			for z in range(0,starttab-1):
				all_sheets.remove(all_sheets[z])
		for page in all_sheets:
			for row in range(0,len(wb[page].column[0])):
				for column in range(0,len(wb[page].row[0])):
					try:
						if colname.lower() == wb[page][row,column].lower():
							header_column = column
							header_row = row
							header_found = True
					except:
						pass
		for page in all_sheets:
			for column in range(0,len(wb[page].row[0])):
				headerrow.append(wb[page][header_row,column])
		for page in all_sheets:
			for row in range(0,len(wb[page].column[0])):
				htmlcode = HTML.table()
				for column in range(0,len(wb[page].row[0])):
					if str(wb[page][row,header_column]) in accountlist:
						accountrow.append(wb[page][row,column])
				if len(accountrow) > 0:
					for cell in range(0,len(accountrow)):
						if ("ssn" in str(headerrow[cell]).lower() or "tax" in str(headerrow[cell]).lower() or "social" in str(headerrow[cell]).lower() or str(headerrow[cell]).lower() == 'tin' or "soc_sec_num" in str(headerrow[cell]).lower() or "ss #" in str(headerrow[cell]).lower()) and (len(str(accountrow[cell])) != "0" or len(str(accountrow[cell])) != "1"):
							accountrow[cell] = 'XXX-XX-X' + str(accountrow[cell][-3:])
						try:
							if isinstance(accountrow[cell],datetime.date) == True:
								accountrow[cell] = accountrow[cell].strftime("%m-%d-%Y")
						except:
							pass
						try:
							if "ph" in str(headerrow[cell]).lower() and isinstance(int(accountrow[cell]),int) == True:
								accountrow[cell] = '({}) {}-{}'.format(accountrow[cell][0:3],accountrow[cell][3:6],accountrow[cell][6:])
						except:
							pass
						if "email" in str(headerrow[cell]).lower():
							accountrow[cell] = 'XXXXX'
						if "sale_price" in str(headerrow[cell]).lower() or "proceeds" in str(headerrow[cell]).lower():
							accountrow[cell] = 'XXXXX'
					if not os.path.exists('/mnt/consentorders/' + str(datetime.date.today().strftime("%m-%d-%Y")) + '/'):
						os.makedirs('/mnt/consentorders/' + str(datetime.date.today().strftime("%m-%d-%Y")) + '/')
					if os.path.exists('/mnt/consentorders/' + str(datetime.date.today().strftime("%m-%d-%Y")) + '/' + str(wb[page][row,header_column]) + '.pdf') == False:
						htmlcode += HTML.table([headerrow,accountrow],border=0,style=(styling))
						pdfkit.from_string(htmlcode, '/mnt/consentorders/' + str(datetime.date.today().strftime("%m-%d-%Y")) + '/' + str(wb[page][row,header_column]) + '.pdf', options={'orientation': 'Landscape', 'quiet': ''})
						found_accounts += 1
					accountrow = []
				self.update_state(state='PROGRESS',meta={'current': found_accounts, 'total': len(accountlist), 'status': 'working...'})
	except:
		pass
	try:
		os.remove(filename)
		os.remove(filename.split(".")[0] + ".pckl")
	except:
		pass
	return {'current': 100, 'total': 100, 'status': 'Task completed!','result': 200}

@app.route('/longtask', methods=['POST'])
def longtask():
	task = long_task.apply_async()
	return jsonify({}), 202, {'Location': url_for('taskstatus', task_id=task.id)}

@app.route('/status/<task_id>')
def taskstatus(task_id):
	task = long_task.AsyncResult(task_id)
	if task.state == 'PENDING':
		response = {
			'state': task.state,
			'current': 0,
			'total': 1,
			'status': 'Pending...'
		}
	elif task.state != 'FAILURE':
		response = {
			'state': task.state,
			'current': task.info.get('current', 0),
			'total': task.info.get('total', 1),
			'status': task.info.get('status', '')
		}
		if 'result' in task.info:
			response['result'] = task.info['result']
	else:
		# something went wrong in the background job
		response = {
			'state': task.state,
			'current': 1,
			'total': 1,
			'status': str(task.info),  # this is the exception raised
		}
	return jsonify(response)

@app.route('/', methods=['GET', 'POST'])
def upload():
	if request.method == 'POST' and 'excel' in request.files:
		file = request.files['excel']
		filename = request.files['excel'].filename
		file.save(filename)
		numbers = request.form.getlist('accounts')
		starttab = request.form.getlist('starttab')[0]
		colname = request.form.getlist('colname')[0]
		accountlist = numbers[0].splitlines()
		accountlist = [i.strip() for i in accountlist]
		with open("{}.pckl".format(filename.split(".")[0]), 'wb') as fp:
			pickle.dump([filename,accountlist,starttab,colname], fp)
	return render_template("upload.html")

@app.route("/checker", methods=['GET', 'POST'])
def checker():
	today = datetime.date.today().strftime("%m-%d-%Y")
	if request.method == 'POST':
		numbers = request.form.getlist('accounts')
		current_date = request.form.getlist('date')
		accountlist = numbers[0].splitlines()
		accountlist = [i.strip() for i in accountlist]
		missing_account = []
		missing_count = 0
		for i in range(0,len(accountlist)):
			if os.path.exists('/mnt/consentorders/' + str(current_date[0]) + '/' + str(accountlist[i]) + '.pdf') == False:
				flash(Markup(str(accountlist[i]).strip()))
				missing_count += 1
		flash(Markup("There was " + str(missing_count) + " missing account(s)"))
		return render_template('checker.html')
	return render_template('checker.html', today=today)

@app.route("/delete", methods=['GET', 'POST'])
def delete():
	today = datetime.date.today().strftime("%m-%d-%Y")
	if request.method == 'POST':
		numbers = request.form.getlist('accounts')
		current_date = request.form.getlist('date')
		accountlist = numbers[0].splitlines()
		accountlist = [i.strip() for i in accountlist]
		delete_account = []
		delete_count = 0
		for i in range(0,len(accountlist)):
			if os.path.exists('/mnt/consentorders/' + str(current_date[0]) + '/' + str(accountlist[i]) + '.pdf') == True:
				os.remove('/mnt/consentorders/' + str(current_date[0]) + '/' + str(accountlist[i]) + '.pdf')
				flash(Markup(str(accountlist[i]).strip()))
				delete_count += 1
		flash(Markup("There was " + str(delete_count) + " account(s) deleted."))
		return render_template('delete.html')
	return render_template('delete.html', today=today)

def is_float(input):
	try:
		num = float(input)
	except ValueError:
		return False
	return True

def is_int(input):
	try:
		num = int(input)
	except ValueError:
		return False
	return True

if __name__ == "__main__":
	# start web server
	app.run(
		#debug=True
		threaded=True,
		host='0.0.0.0',
		port=80
	)
