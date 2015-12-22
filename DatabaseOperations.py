__author__ = 'Kaiqun'

import pyodbc
import os
import datetime

# THis file inserts database records from excel sheets

def Datafetchor():
	currentPath = os.path.dirname(os.path.realpath(__file__))

	with open(currentPath + '/DBCredentials/MSSQL', 'r') as crdts:
		CrStr = crdts.readline()
		DBDriver = CrStr.split(',')[0]
		DBServer = CrStr.split(',')[1]
		DBDatabase = CrStr.split(',')[2]
		DBUid = CrStr.split(',')[3]
		DBPwd = CrStr.split(',')[4]

	connStr = 'DRIVER={%s};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s' % (DBDriver, DBServer, DBDatabase, DBUid, DBPwd)

	cnxn = pyodbc.connect(connStr)
	cursor = cnxn.cursor()

	try:
		InsertCommand = "EXEC RTMS_Last_Date_Grouped"
		cursor.execute(InsertCommand)
		ReturningList = []
		count = 0
		while True:
			tmpList = []
			row = cursor.fetchone()
			if not row:
				break
			count += 1
			tmpList.append(count)
			tmpList.append(row.DTime.strftime('%Y-%m-%d %H:%M:%S'))
			tmpList.append(row.NewNames.strip())
			ReturningList.append(tmpList)
		return ReturningList
	except Exception, e:
		print 'Some Error'
		return None