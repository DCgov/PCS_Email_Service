from os import listdir
from os.path import isfile, join, isdir
import datetime
import HTML

LocationsMapping = {
	"1": "Interstate 395 Bridges",
	"2": "Interstate 395 Bridges",
	"3": "Interstate 395 Bridges",
	"4": "Theodore Roosevelt Bridge (I-66)",
	"5": "Theodore Roosevelt Bridge (I-66)",
	"6": "Key Bridge (US Route 29)",
	"7": "Key Bridge (US Route 29)",
	"8": "Douglas Bridge (South Capitol Street)",
	"9": "Whitney Young Memorial Bridge (East Capitol Street)",
	"10": "Rhode Island Ave, NE, West of Eastern Ave",
	"11": "New Hampshire Ave, NE, at B&O Overpass",
	"12": "Georgia Ave, NW, between Eastern Ave and Alaska Ave",
	"13": "Georgia Ave, NW, between Eastern Ave and Alaska Ave",
	"14": "Connecticut Ave, NW, between Chevy Chase Circle and Oliver St",
	"15": "Nebraska Ave, NW, South of Ward Circle",
	"16": "Massachusetts Ave, NW, East of 30th St",
	"17": "Massachusetts Ave, NW, East of 30th St",
	"18": "K St NW, between 18th St and 19th St",
	"19": "K St NW, between 18th St and 19th St",
	"20": "16th St NW, North of M St",
	"21": "13th St NW, between Eye St and K St",
	"22": "13th St NW, between Eye St and K St",
	"29": "Pennsylvania Ave, SE, between 40th St and Fort Davis St",
	"23": "Alabama Ave, SE between Hartford St and Gainesville St",
	"24": "Wisconsin Ave NW, at DC Line",
	"25": "16th St NW, at DC Line",
	"26": "Interstate 395 NW, South of New York Ave",
	"27": "Kenilworth Ave, NE at DC Line",
	"28": "Kenilworth Ave, NE at DC Line",
	"30": "11th St Bridge",
	"31": "11th St Bridge"
}


def buildHTMLcell(tilDate):
	realdatetime = datetime.datetime.strptime(tilDate, '%Y-%m-%d')
	if (datetime.datetime.now() - realdatetime) < datetime.timedelta(days = 5):
		result = HTML.TableCell(tilDate, bgcolor = 'lightgreen')
	elif (datetime.datetime.now() - realdatetime) < datetime.timedelta(days = 10):
		result = HTML.TableCell(tilDate, bgcolor = 'yellow')
	else:
		result = HTML.TableCell(tilDate, bgcolor = 'pink')
	return result


def buildHTMLtable(inputList):
	t = HTML.Table(
		header_row = ['StationID (ADR)', 'Last Available Count', 'Location'],
		col_styles = ['background-color:lightgray; font-weight:bold; text-align:center', 'text-align:center',
		              'text-align:center'])

	for row in inputList:
		table_row = []
		table_row.append(HTML.TableCell(row.split('|')[0]))
		table_row.append(buildHTMLcell(row.split('|')[1]))
		table_row.append(HTML.TableCell(row.split('|')[2]))

		t.rows.append(table_row)

	# for row in tally_list:
	# # create the emtpy row
	#     if row[0] == '-':
	#         #this is the blank one
	#         continue

	#     table_row = []
	#     #roll through each sub list in the tally list and inspect each value
	#     for cell in row:
	#         this_cell = buildHTMLcell(cell)
	#         table_row.append(this_cell)

	#     #append the summary info to the end of this row
	#     if row[0] in summary:
	#         goods = float(summary[row[0]]['good'])
	#         warnings = float(summary[row[0]]['warning'])
	#         bads = float(summary[row[0]]['bad'])

	#         total = goods + warnings + bads
	#         table_row.append(HTML.TableCell("{0:.2f}%".format(goods / total * 100)))
	#         table_row.append(HTML.TableCell("{0:.2f}%".format(warnings / total * 100)))
	#         table_row.append(HTML.TableCell("{0:.2f}%".format(bads / total * 100)))


	#     # append the row with all cells:
	#     t.rows.append(table_row)

	return t


TargetYear = datetime.datetime.now().strftime('%Y')

# targetDir = 'C:\\ProgramData\\Peek Traffic\\Viper\\Output1\\Binned Data Reports\\Class\\' + TargetYear
targetDir = 'C:\\ProgramData\\Peek Traffic\\Viper V1.5.8\\Output1\\Binned Data Reports\\Class\\' + TargetYear
# targetDir = 'C:\Users\Kaiqun\Desktop'

writingFile = open(
	'\\\\sddotfile1\\toa\\ITS_Design_Implementation\\Data Connect\\reports\\Loops\\' + datetime.datetime.now().strftime(
		'%Y-%m-%d.html'), 'w+')

while not isdir(targetDir):
	print 'Year Selection Error. Please select again'
	TargetYear = str(raw_input('Please input the Year: '))
	targetDir = 'C:\\ProgramData\\Peek Traffic\\Viper\\Output1\\Binned Data Reports\\Class\\' + TargetYear

onlydirs = [targetDir + '\\' + name for name in listdir(targetDir) if isdir(join(targetDir, name))]

tmpListing = []
for folder in onlydirs:
	onlyfiles = [TargetYear + '-' + f.split('_')[-1:][0][:2] + '-' + f.split('_')[-1:][0][2:4] for f in listdir(folder)
	             if isfile(join(folder, f))]
	tmpListing.append(str(int(folder.split('\\')[-1:][0])) + '|' + '|'.join(onlyfiles[-1:]) + '|' + LocationsMapping[
		str(int(folder.split('\\')[-1:][0]))])

writingFile.write(str(buildHTMLtable(tmpListing)))