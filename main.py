from openpyxl import load_workbook
from datetime import date
import time
import sys

# region startup config
today, note = '', ''
hours, stop_time = 0, 0
today_already_has_entry = False

file_name = "Working Hours.xlsx"  # Change filename or use path if needed
sheet_name = "Sheet 1"  # Change sheet name if needed

wb = load_workbook(filename=file_name)
sheet_ranges = wb[sheet_name]
# endregion


def get_today():
	today_data = str(date.today()).split('-')
	return f"{today_data[2]}.{today_data[1]}.{today_data[0]}"


def find_row():
	global today_already_has_entry
	for row_num in range(2, 999):
		if sheet_ranges['A' + str(row_num)].value == today:
			today_already_has_entry = True
			return str(row_num)
		elif sheet_ranges['A' + str(row_num)].value is None:
			return str(row_num)


def calculate_working_time(start, end):  # returns working time in hours
	working_time = end - start - stop_time
	return round(working_time / 60 / 60, 2)


def upload_data():
	row_num = find_row()
	if today_already_has_entry:
		sheet_ranges[f'D{row_num}'] = sheet_ranges['D'+str(row_num)].value + hours
		sheet_ranges[f'E{row_num}'] = sheet_ranges['E'+str(row_num)].value + " | " + note
	else:
		sheet_ranges[f'A{row_num}'] = today
		sheet_ranges[f'D{row_num}'] = hours
		sheet_ranges[f'E{row_num}'] = note


def save_file():
	print("Saving file")

	try:
		wb.save(filename=file_name)
	except PermissionError:
		input("Saving failed. Please, make sure file is not opened in another app, then press enter to try again.")
		try:
			wb.save(filename=file_name)
		except PermissionError:
			print("Saving failed again, please save your data manually.")
			print(f"Hours worked: {hours} (= {int(hours)}h and {(hours-int(hours))*60}m)")
			print(f"Note: {note}")

	print("File saved successfully")


def session_stopped():
	global stop_time
	stop_start = time.time()
	print("TIME STOPPED")
	command = input("Type 'continue' or 'end': ")
	while command not in ['continue', 'end']:
		command = input("Incorrect. Type 'continue' or 'stop': ")

	if command == 'continue':
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
	print("TIME ENDED")
	hours = calculate_working_time(work_start, work_end)
	note = input("Enter note: ")
	upload_data()
	save_file()
	input("Press enter to exit")
	sys.exit()


main()
