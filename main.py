from exceler import Exceler
import webbrowser, os


if __name__ == '__main__':
	workbook_name = 'finances.xlsx'
	exceler = Exceler(workbook_name)
	exceler.read_workbook_to_python()
	exceler.write_workbook_to_excel()
	webbrowser.open('file://' + os.path.realpath(workbook_name))



