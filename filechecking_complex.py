import pyodbc
import datetime
import sys, os
import csv
import HTML
import re
import xlrd

from os import listdir
from os.path import isfile, join, isdir

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

# http://stackoverflow.com/questions/32404/is-it-possible-to-run-a-python-script-as-a-service-in-windows-if-possible-how


# function lifted from http://stackoverflow.com/questions/5967500/how-to-correctly-sort-a-string-with-a-number-inside
def atoi(text):
	return int(text) if text.isdigit() else text


# function lifted from http://stackoverflow.com/questions/5967500/how-to-correctly-sort-a-string-with-a-number-inside
def natural_keys(text):
	'''
    alist.sort(key=natural_keys) sorts in human order
    http://nedbatchelder.com/blog/200712/human_sorting.html
    '''
	return [atoi(c) for c in re.split('(\d+)', text)]


def checkistime(ipnutlist):
	try:
		reslt = ipnutlist[1]
		return True
	except IndexError:
		return False


def getExpectedCharacteristics(mycsv):
	'''
    this loads all of the known counters (station ids, number of channels)
    into 1 of four types of dictionaries.  All 4 dictionaries are returned
    '''
	loops = dict()
	microwaves = dict()
	infrareds = dict()
	videos = dict()

	input_file = csv.DictReader(open(mycsv))

	for row in input_file:
		if row["type"] == "loop":
			loops[row["StationID"] + "_" + row["Channel"]] = {"expected_daily_counts": 24, "expected_smallest_count": 1,
			                                                  "counts_this_period": 0, "invalid_counts_this_period": []}

		if row["type"] == "microwave":
			microwaves[row["Microwave"] + "_" + row["Channel"]] = {"expected_daily_counts": 24}

		if row["type"] == "infrared":
			infrareds[row["Pair"]] = {"expected_daily_counts": 24}

		if row["type"] == "Video":
			videos[row["Zoneid"]] = {"expected_daily_counts": 24}

	return loops, microwaves, infrareds, videos


def evaluateResults(resultset, todays_date):
	invalidcountlist = []
	totalcountlist = []

	for k, v in resultset.items():

		if v['counts_this_period'] != v['expected_daily_counts']:
			# counts not adding up.  flag
			flag = "Station " + k.split("_")[0] + ", Channel, " + k.split("_")[
				1] + ", " + todays_date + " :  Expected Total Daily Counts - " + str(
				v['expected_daily_counts']) + ".  Non-zero Counts Today - " + str(v['counts_this_period'])
			totalcountlist.append(flag)
		if len(v['invalid_counts_this_period']) > 0:
			# some low-value or invalid counts.  flag
			for invalid in v['invalid_counts_this_period']:
				invalidcountlist.append(invalid)

	return totalcountlist, invalidcountlist


def buildTallyTable(loops):
	html_table = []
	station_row = []
	stations = loops.keys()
	stations.sort(key = natural_keys)
	previous_station = 1

	for thisstation in stations:
		stationid = thisstation.split("_")[0]
		channelid = thisstation.split("_")[1]

		if previous_station == stationid:
			# we are still on the same station, continue to build this station row
			station_row.append(loops[thisstation]['counts_this_period'])

		else:
			# we are starting a new station. Store current data, then clear row and begin anew.
			if len(station_row) < 9:
				station_row.extend('-' * (9 - len(
					station_row)))  #this is to fill in an '-' into the extra fields for those stations that don't have 8 channels
			#station_row.pop(0) #first row is empty.  get rid of it prior to loading.
			html_table.append(station_row)
			station_row = []
			station_row.append(stationid)  #first column is the station id number
			station_row.append(loops[thisstation]['counts_this_period'])

		previous_station = stationid

	# add last loop row.
	if len(station_row) < 9:
		station_row.extend('-' * (9 - len(
			station_row)))  #this is to fill in an '-' into the extra fields for those stations that don't have 8 channels
	#station_row.pop(0) #first row is empty.  get rid of it prior to loading.
	html_table.append(station_row)

	return html_table


def buildHTMLcell(cell):
	if cell == '-':
		# no channel for this station
		result = HTML.TableCell(str(cell), bgcolor = 'lightgray')
	elif cell < 18:
		# missing many counts
		result = HTML.TableCell(str(cell), bgcolor = 'pink')
	elif 18 <= cell < 24:
		# low/missing count
		result = HTML.TableCell(str(cell), bgcolor = 'yellow')
	elif cell == 24:
		# no missing counts
		result = HTML.TableCell(str(cell), bgcolor = 'lightgreen')
	else:
		# count tally good
		result = HTML.TableCell(str(cell), bgcolor = 'lightpurple')

	return result


def buildHTMLcell_sourcefile(tilDate):
	realdatetime = datetime.datetime.strptime(tilDate, '%Y-%m-%d')
	if (datetime.datetime.now() - realdatetime) < datetime.timedelta(days = 5):
		result = HTML.TableCell(tilDate, bgcolor = 'lightgreen')
	elif (datetime.datetime.now() - realdatetime) < datetime.timedelta(days = 10):
		result = HTML.TableCell(tilDate, bgcolor = 'yellow')
	else:
		result = HTML.TableCell(tilDate, bgcolor = 'pink')
	return result


def onefileoperate(filedir):
	currentYear = datetime.datetime.now().strftime('%Y')
	currentDateTime = ''

	onefile = filedir.split('\\')[-1:][0]
	currentDate = currentYear + '-' + onefile.split('_')[-1:][0][:2] + '-' + onefile.split('_')[-1:][0][2:4]
	StationID = onefile.split('_')[-2:][0]

	workbook = xlrd.open_workbook(filedir)

	InsertionHandle = []

	if 'Sheet1' == workbook.sheet_names()[0]:
		worksheet = workbook.sheet_by_name('Sheet1')
		num_rows = worksheet.nrows - 1

		curr_row = 3
		channelIndex = 0

		while curr_row < num_rows:
			curr_row += 1

			# Break loop if is at bottom
			if worksheet.cell_value(curr_row, 1) == 'Total':
				break

			# Skip Class heading row and Totals
			if worksheet.cell_type(curr_row, 1) == 0 and worksheet.cell_type(curr_row, 2) == 0:
				continue
			if worksheet.cell_value(curr_row, 2) == 'Total':
				continue

			# Counting Channels
			if worksheet.cell_type(curr_row, 1) == 1 and worksheet.cell_value(curr_row, 2) == 'Channel 1':
				currentDateTime = currentDate + ' ' + "%02d" % (
					int(worksheet.cell_value(curr_row, 1).split(':')[0]) - 1) + ':00:00'
				channelIndex = 1
			else:
				channelIndex += 1

			# print str(int(StationID)) + '--' + currentDateTime
			countList = [int(worksheet.cell_value(curr_row, i)) for i in range(3, 20) if
			             worksheet.cell_value(curr_row, i) != '']

			InsertionHandle.append([int(StationID), currentDateTime, channelIndex, countList[1]])

	elif 'Sheet' == workbook.sheet_names()[0]:
		worksheet = workbook.sheet_by_name('Sheet')
		num_rows = worksheet.nrows - 1

		curr_row = 7

		while curr_row < num_rows:
			if worksheet.cell_value(curr_row, 14) == 'All Lanes':
				break

			if worksheet.cell_type(curr_row, 0) == 3 and worksheet.cell_type(curr_row, 14) == 1:
				channelIndex = int(worksheet.cell_value(curr_row, 14)[-2:])

				curr_row += 3
				while checkistime(worksheet.cell_value(curr_row, 0).split(':')):
					currentDateTime = currentDate + ' ' + worksheet.cell_value(curr_row, 0) + ':00'
					countList = [int(worksheet.cell_value(curr_row, i)) for i in range(17) if
					             worksheet.cell_type(curr_row, i) == 2]
					# print str(int(StationID)) + '--' + str(channelIndex) + '--' + currentDateTime
					InsertionHandle.append([int(StationID), currentDateTime, channelIndex, countList[1]])
					curr_row += 1

			curr_row += 1

	else:
		print 'No Sheet Name Found!'

	return InsertionHandle


def channelStast(inputStationIDList):
	returningList = []
	for oneStation in inputStationIDList:
		tmpList = []
		tmpList.append(int(oneStation.split('\\')[8]))
		Channellist = onefileoperate(oneStation)
		Cindex = 1
		while Cindex <= 8:
			allTimeinChannel = [onetime for onetime in Channellist if onetime[2] == Cindex]
			if allTimeinChannel == []:
				tmpList.append('-')
			else:
				count = 0
				for aline in allTimeinChannel:
					# Class #2 is less than
					if aline[3] < 2:
						count += 0
					else:
						count += 1
				tmpList.append(count)
			Cindex += 1
		returningList.append(tmpList)
	return returningList


def buildHTMLtable(tally_list, summary, sourcefileList, chennelChecks):
	t = HTML.Table(
		header_row = ['StationID (ADR)', 'Channel1', 'Channel2', 'Channel3', 'Channel4', 'Channel5', 'Channel6',
		              'Channel7',
		              'Channel8', 'Last Available Count', 'Location', 'Good %(Last 30 Days)', 'Warning %(Last 30 Days)',
		              'Bad %(Last 30 Days)'],
		col_styles = ['background-color:lightgray; font-weight:bold; text-align:center', 'text-align:center',
		              'text-align:center', 'text-align:center', 'text-align:center', 'text-align:center',
		              'text-align:center', 'text-align:center', 'text-align:center', 'text-align:center',
		              'text-align:center',
		              'background-color:lightgray; font-weight:bold; text-align:center',
		              'background-color:lightgray; font-weight:bold; text-align:center',
		              'background-color:lightgray; font-weight:bold; text-align:center'])
	sourceDict = {}
	for onerow in sourcefileList:
		sourceDict.update({onerow.split('|')[0]: onerow.split('|')[1:]})

	for row in chennelChecks:
		# create the emtpy row
		if row[0] == '-':
			# this is the blank one
			continue

		table_row = []
		# roll through each sub list in the tally list and inspect each value
		count = 0
		for cell in row:
			if count == 0:
				table_row.append(HTML.TableCell(row[count]))
			else:
				this_cell = buildHTMLcell(cell)
				table_row.append(this_cell)
			count += 1

		if str(row[0]) == '25':
			table_row.append(HTML.TableCell('-'))
			table_row.append(HTML.TableCell('-'))

		if str(row[0]) in sourceDict:
			sourceDict[str(row[0])]
			table_row.append(buildHTMLcell_sourcefile(sourceDict[str(row[0])][0]))
			table_row.append(HTML.TableCell(sourceDict[str(row[0])][1]))

		#append the summary info to the end of this row
		if str(row[0]) in summary:
			goods = float(summary[str(row[0])]['good'])
			warnings = float(summary[str(row[0])]['warning'])
			bads = float(summary[str(row[0])]['bad'])

			total = goods + warnings + bads
			table_row.append(HTML.TableCell("{0:.2f}%".format(goods / total * 100)))
			table_row.append(HTML.TableCell("{0:.2f}%".format(warnings / total * 100)))
			table_row.append(HTML.TableCell("{0:.2f}%".format(bads / total * 100)))

		# append the row with all cells:
		t.rows.append(table_row)

	return t


def checkLoops_Daily(thisdate = datetime.datetime.now(), ishist = False, loops = None):
	"""
    This function evaluates the LoopsClass table.  Counts are stored every hour
    and we expect to see 1 or more in the 'Count' column for each channel.  If there
    is a 0 in a Class 2 count, flag.
    """

	cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.40.62.167;DATABASE=speed;UID=tirs;PWD=keepsafe')

	cursor = cnxn.cursor()

	if ishist:
		yesterday = thisdate
	else:
		yesterday = thisdate - datetime.timedelta(days = 3)

	st_dt = yesterday.replace(hour = 0, minute = 0, second = 0, microsecond = 0)
	end_dt = yesterday.replace(hour = 23, minute = 59, second = 59, microsecond = 0)
	todays_date = yesterday.strftime('%m/%d/%Y')
	todays_date_time = yesterday.strftime('%m/%d/%Y at %I%p')

	# SQL statement
	sql = """Select DISTINCT StationID, DateTime, Channel, Class, Count From [speed].[tirs].[LoopsClass_new] Where DateTime Between '%s' And '%s' Order By StationID, DateTime""" % (
		st_dt, end_dt)
	cursor.execute(sql)

	row = cursor.fetchall()
	count = 0
	for record in row:
		this_station = record.StationID.lstrip("0")

		this_station_id = this_station + "_" + str(record.Channel)

		#we only care about passenger vehicles:
		if record.Class == 2:
			#check to see if this station/channel is in our key.csv loops dict
			if this_station_id in loops:
				#check to see if the count on this station/channel is at or above minimum:
				if loops[this_station_id]["expected_smallest_count"] <= record.Count:
					#increment counter to indicate that this was a valid count
					loops[this_station_id]["counts_this_period"] += 1
				else:
					#count for this class is not acceptable. store for report flag
					this_invalid_record = "Station " + this_station_id.split("_")[0] + ", Channel, " + \
					                      this_station_id.split("_")[
						                      1] + ", " + todays_date_time + ", Low Count: " + str(record.Count)
					#loops[this_station_id]["invalid_counts_this_period"].append(this_invalid_record)

	cnxn.close()
	detailed_list = evaluateResults(loops, todays_date)

	return detailed_list


def entry():
	# simudate = datetime.datetime.today() - datetime.timedelta(days=15)
	# while True:
	# simudate += datetime.timedelta(days=1)
	scriptloc = sys.path[0]
	csvpath = os.path.join(sys.path[0], "key.csv")

	clipdays = 3  # this number will adjust the summary period.  Think of this as the 'begin summary X days prior'

	# Get historical tally first
	base = datetime.datetime.today() - datetime.timedelta(days = 3)
	# base = simudate - datetime.timedelta(days=3)
	# base += datetime.timedelta(days=1)
	date_list = [base - datetime.timedelta(days = x) for x in range(0, 31 - clipdays)]
	historicaldatelist = date_list[clipdays:]  #remove the three most recent days.

	summary = dict()

	for i in range(32):
		summary[str(i)] = {'good': 0, 'warning': 0, 'bad': 0}

	for adate in historicaldatelist:
		loops, microwaves, infrareds, videos = getExpectedCharacteristics(csvpath)
		loop_results_hist = checkLoops_Daily(adate, True,
		                                     loops)  #add date as parameter to run a specific date like datetime.datetime(2014, 5, 30)
		tally_list_hist = buildTallyTable(loops)

		# print adate
		for station in tally_list_hist:

			if station[0] == '-':
				pass

			else:
				if station[0] == '28':
					thisdate = adate.strftime('%m/%d/%Y')
					with open(os.path.join(scriptloc, "loop_station_" + station[0] + ".txt"), "a") as f:
						f.write(",".join(map(str, station)) + "," + thisdate + "\n")

				for channel in station[1:]:

					if channel == '-':
						#no channel for this station
						pass
					elif channel < 18:
						#missing many counts
						summary[station[0]]['bad'] += 1
					elif 18 <= channel < 24:
						#low/missing count
						summary[station[0]]['warning'] += 1
					elif channel == 24:
						#no missing counts
						summary[station[0]]['good'] += 1
					else:
						print "unexpected error"

	#Perform Daily Summary Last
	loops, microwaves, infrareds, videos = getExpectedCharacteristics(csvpath)

	#Get most recent count
	loop_results = checkLoops_Daily(thisdate = base, ishist = True,
	                                loops = loops)  #add date as parameter to run a specific date like datetime.datetime(2014, 5, 30)
	tally_list = buildTallyTable(loops)

	TargetYear = datetime.datetime.now().strftime('%Y')
	targetDir = 'C:\\ProgramData\\Peek Traffic\\Viper V1.5.8\\Output1\\Binned Data Reports\\Class\\' + TargetYear
	writingFile = open(
		'\\\\sddotfile1\\toa\\ITS_Design_Implementation\\Data Connect\\complex\\Loops\\' + datetime.datetime.now().strftime(
			'%Y-%m-%d.html'), 'w+')
	onlydirs = [targetDir + '\\' + name for name in listdir(targetDir) if isdir(join(targetDir, name))]
	tmpListing = []

	CheckFileList = []

	for folder in onlydirs:
		onlyfiles = [TargetYear + '-' + f.split('_')[-1:][0][:2] + '-' + f.split('_')[-1:][0][2:4] for f in
		             listdir(folder) if isfile(join(folder, f))]
		tmpListing.append(
			str(int(folder.split('\\')[-1:][0])) + '|' + '|'.join(onlyfiles[-1:]) + '|' + LocationsMapping[
				str(int(folder.split('\\')[-1:][0]))])

		os.chdir(folder)
		# TFiles = [folder + '\\' + file for file in glob.glob("*.xlsx") if file.split('_')[-1:][0][:4] == (onlyfiles[-1:][0].split('-')[1] + onlyfiles[-1:][0].split('-')[2])]
		TFiles = [folder + '\\' + file for file in os.listdir(folder) if re.match(
			'.*' + (onlyfiles[-1:][0].split('-')[1] + onlyfiles[-1:][0].split('-')[2]) + '\d{4}\.xlsx|.*' + (
				onlyfiles[-1:][0].split('-')[1] + onlyfiles[-1:][0].split('-')[2]) + '\d{4}\.xls', file)]
		CheckFileList.append(TFiles[0])

	html_table = buildHTMLtable(tally_list, summary, tmpListing, channelStast(CheckFileList))
	text_table = str(html_table)

	writingFile.write(text_table.strip().replace('\n', '').replace('\r', ''))

	# MailSend(text_table, base)
	print 'Complex Generated! '


entry()