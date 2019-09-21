from exceler import Exceler
import webbrowser, os, time


if __name__ == '__main__':
	workbook_name = 'finances.xlsx'
	five_hours = 60*60*5
	try:
		while True:
			exceler = Exceler(workbook_name)
			exceler.read_workbook_to_python()
			exceler.write_workbook_to_excel()
			time.sleep(five_hours)
	except:	# opens your excel file when you kill the program	
		webbrowser.open('file://' + os.path.realpath(workbook_name))



