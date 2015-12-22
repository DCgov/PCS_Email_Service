from os import listdir
from os.path import isfile, join, isdir
import datetime
import HTML
from DatabaseOperations import Datafetchor


def buildHTMLcell(tilDate):
	realdatetime = datetime.datetime.strptime(tilDate, '%Y-%m-%d %H:%M:%S')
	if (datetime.datetime.now() - realdatetime) < datetime.timedelta(days = 5):
		result = HTML.TableCell(tilDate, bgcolor = 'lightgreen')
	elif (datetime.datetime.now() - realdatetime) < datetime.timedelta(days = 10):
		result = HTML.TableCell(tilDate, bgcolor = 'yellow')
	else:
		result = HTML.TableCell(tilDate, bgcolor = 'pink')
	return result


def buildHTMLtable(inputList):
	t = HTML.Table(
		header_row = ['StationID', 'Last Available Count', 'Location'],
		col_styles = ['background-color:lightgray; font-weight:bold; text-align:center', 'text-align:center',
		              'text-align:center'])

	for row in inputList:
		table_row = []
		table_row.append(HTML.TableCell(row[0]))
		table_row.append(buildHTMLcell(row[1]))
		table_row.append(HTML.TableCell(row[2]))

		t.rows.append(table_row)

	return t


writingFile = open(
	'\\\\sddotfile1\\toa\\ITS_Design_Implementation\\Data Connect\\reports\\RTMS\\' + datetime.datetime.now().strftime(
		'%Y-%m-%d.html'), 'w+')

writingFile.write(str(buildHTMLtable(Datafetchor())))