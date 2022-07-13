from openpyxl import load_workbook
from datetime import date
import time
import sys
import os

# region startup config
today, note = '', ''
hours, stop_time = 0, 0
today_already_has_entry = False

file_name = "Working Hours.xlsx"  # Change filename or use path if needed
sheet_name = "Sheet 1"  # Change sheet name if needed

wb = load_workbook(filename=file_name)
sheet_ranges = wb[sheet_name]
# endregion


def clear():  # clears console
	os.system('cls')


def get_today():  # returns today's date in a format dd.mm.yyyy
	today_data = str(date.today()).split('-')
	return f"{today_data[2]}.{today_data[1]}.{today_data[0]}"


def find_row():  # returns row's number that will be used to save new data
	global today_already_has_entry
	for row_num in range(2, 999):
		if sheet_ranges['A' + str(row_num)].value == today:
			today_already_has_entry = True
			return str(row_num)
		elif sheet_ranges['A' + str(row_num)].value is None:
			return str(row_num)


def calculate_working_time(start, end):  # returns working time in hours
	working_time = round((end - start - stop_time) / 60 / 60, 2)
	working_hours = int(working_time)
	working_minutes = (working_time - int(working_time)) * 60

	if working_hours == 0:
		print(f"You were working for {working_minutes} minutes")
	elif working_minutes == 0:
		print(f"You were working for {working_hours} hours")
	elif working_minutes == 1:
		print(f"You were working for {working_hours} hours and 1 minute")
	else:
		print(f"You were working for {working_hours} hours and {working_minutes} minutes")

	return working_time


def upload_data():  # uploads data to .xlsx file
	row_num = find_row()
	if today_already_has_entry:
		sheet_ranges[f'D{row_num}'] = sheet_ranges['D'+str(row_num)].value + hours
		if note != '':
			sheet_ranges[f'E{row_num}'] = sheet_ranges['E'+str(row_num)].value + " | " + note
	else:
		sheet_ranges[f'A{row_num}'] = today
		sheet_ranges[f'D{row_num}'] = hours
		sheet_ranges[f'E{row_num}'] = note


def save_file():  # saves .xlsx file
	print("Saving file")

	try:
		wb.save(filename=file_name)
		print("File saved successfully")
	except PermissionError:
		input("Saving failed. Please, make sure file is not opened in another app, then press enter to try again.")
		try:
			wb.save(filename=file_name)
			print("File saved successfully")
		except PermissionError:
			print("Saving failed again, please save your data manually.")


def session_stopped():
	global stop_time
	stop_start = time.time()
	clear()
	print("TIME STOPPED")
	command = input("Type 'continue' or 'end': ")
	while command not in ['continue', 'end']:
		command = input("Incorrect. Type 'continue' or 'stop': ")

	if command == 'continue':
		clear()
		print("TIME STARTED AGAIN")
		stop_end = time.time()
		stop_time += stop_end - stop_start
		session_running()
	else:
		stop_end = time.time()
		stop_time += stop_end - stop_start


def session_running():
	command = input("Type 'stop' or 'end': ")
	while command not in ['stop', 'end']:
		command = input("Incorrect. Type 'end' or 'stop': ")
	if command == 'stop':
		session_stopped()


def main():
	global today, hours, note
	work_start = time.time()
	print("TIME STARTED")
	today = get_today()
	session_running()
	work_end = time.time()
	clear()
	print("TIME ENDED")

	hours = calculate_working_time(work_start, work_end)
	note = input("Enter note: ")
	upload_data()
	save_file()
	input("Press enter to exit")
	sys.exit()


main()
