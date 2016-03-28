import csv, datetime 

# import arcpy
# DetailedSchedule = arcpy.GetParameterAsText(0)
# Abstracts = arcpy.GetParameterAsText(1)
# Posters = arcpy.GetParameterAsText(2)
# ScheduleGroup = arcpy.GetParameterAsText(3)
# outWordFile = arcpy.GetParameterAsText(4)
# outWebsiteFile = arcpy.GetParameterAsText(5)

class HardCodedParameters(object):
	def __init__(self):
		self.DetailedSchedule = "Detailed Schedule.csv"
		self.Abstracts =  "Abstracts.csv"
		self.Posters =  "Abstracts Organization - Moderators.csv"
		self.outWordFile =  "outWord.html"
		self.outWebsiteFile =  "outWeb.html"
		self.ScheduleGroups = "ScheduleGroups.csv"

params = HardCodedParameters()

def scheduleGroupsToDict(filename):
	csvfile = open(filename, 'r')
	csvreader = csv.reader(csvfile)
	printedGroupNames = {(row[0], row[1]): row[2] for row in csvreader}
	return printedGroupNames

printedGroupNames = scheduleGroupsToDict(params.ScheduleGroups)

# make list of dictionaries, associating each piece of data
# with its column name
def readCSV(fileName):
	csvfile = open(fileName, 'r')
	csvreader = csv.reader(csvfile)
	rows = list(csvreader)
	headerRow = rows[0]
	dataRows = rows[1:]
	return dataRows, headerRow
def makeRowDicts(dataRows, headerRow):
	listOfRowDicts = []
	for row in dataRows:
		# rowDict = {colName : value for (colName, value) in zip(headerRow, row)}
		# rowDict = dict(zip(headerRow, row))
		rowDict = {headerRow[x]:row[x] for x in range(0,len(row))}
		listOfRowDicts.append(rowDict)
	# return [dict(zip(headerRow, row)) for row in dataRows]
	return listOfRowDicts
def readCSVWithHeader(filename):
	x, y = readCSV(filename)
	z = makeRowDicts(x, y)
	return z

schedule = readCSVWithHeader(params.DetailedSchedule)
abst = readCSVWithHeader(params.Abstracts)
poster = readCSVWithHeader(params.Posters)

def correctDayNames(schedule):
	fullDayNames = {'Mon': 'Monday', 'Sun': 'Sunday', 'Tues': 'Tuesday', 'Wed': 'Wednesday'}
	for row in schedule:
		row['Day'] = fullDayNames[row['Day']]


correctDayNames(schedule)

# adds authors and title into the schedule
def associateAbstWithSchedule(abstracts, schedule):
	abstractsById = {entry['No']:entry for entry in abstracts}
	scheduleWithAbstracts = []
	for scheduleRow in schedule:
		copiedRow = {k:v for (k,v) in scheduleRow.items()}
		absId = scheduleRow["Abst"]
		if absId != "-" and absId in abstractsById:
			abstractInfo = abstractsById[absId]
			copiedRow["Author"] = abstractInfo["Authors"]
			copiedRow["Title"] = abstractInfo["Title"]
		scheduleWithAbstracts.append(copiedRow)
	return scheduleWithAbstracts

scheduleWithAbstracts = associateAbstWithSchedule(abst, schedule)

class Poster(object):
	def __init__(self, rawRow):
		self._raw = rawRow
	def title(self):
		return self._raw["Title"]
	def number(self):
		return self._raw["No"]
	def authors(self):
		return self._raw["Authors"]

# adds poster info to schedule
def associatePosterWithSchedule(posterData):
	posters = [row for row in posterData if row["Poster"] == "1"]
	postersBySession = dict()
	for poster in posters:
		sessionName = poster["Session"]
		if sessionName not in postersBySession:
			postersBySession[sessionName] = []
		postersBySession[sessionName].append(Poster(poster))
	return postersBySession

postersBySession = associatePosterWithSchedule(poster)

class Row(object):
	def __init__(self, rawRow):
		self._raw = rawRow
	def startTime(self):
		return self._raw["TimeStart"]
	def endTime(self):
		return self._raw["TimeEnd"]
	def room(self):
		return self._raw["Room"]
	def isTalk(self):
		return self._raw["Notes"] != "Q&A" and self._raw["Abst"] not in ["-", ""]
	def hasTitleAndAuthor(self):
		return "Title" in self._raw
	def titleAndAuthor(self):
		title = self._raw["Title"] + " (" + self._raw["Abst"] + ")"
		author = self._raw["Author"]
		return (title, author)

class Block(object):
	def __init__(self, rows):
		self.rows = rows
	def name(self):
		return self.rows[0]._raw["BlockName"]
	def sessionName(self):
		return self.rows[0]._raw["Block"]
	def startTime(self):
		return self.rows[0].startTime()
	def endTime(self):
		return self.rows[-1].endTime()
	def room(self):
		return self.rows[0].room()
	def hasTalks(self):
		return not self.rows[0]._raw["Abst"] in ["-",""]

class Group(object):
	def __init__(self, blocks):
		self.blocks = blocks
	def name(self):
		return self.blocks[0].rows[0]._raw["Group"]

class Day(object):
	def __init__(self, groups):
		self.groups = groups
	def name(self):
		return self.groups[0].blocks[0].rows[0]._raw["Day"]

def groupConsecutive(items, makeGroupId, 
	groupWrapper = lambda x: x,
	itemWrapper = lambda x: x):
	current = None
	res = []
	for x in items:
		gId = makeGroupId(x)
		assert(gId != None)
		if current == None or gId != current:
			res.append([itemWrapper(x)])
			current = gId
		else:
			# this is the group we are currently building
			res[-1].append(itemWrapper(x))
	return [groupWrapper(g) for g in res]

def associateScheduleDayGroupBlock(scheduleWithAbstracts):
	
	blocks = groupConsecutive(
		items = scheduleWithAbstracts,
		makeGroupId = lambda x: (x["Block"], x["Day"]),
		itemWrapper = Row,
		groupWrapper = Block)

	groups = groupConsecutive(
		items = blocks,
		makeGroupId = lambda x: x.rows[0]._raw["Group"],
		groupWrapper = Group)

	days = groupConsecutive(
		items = groups,
		makeGroupId = lambda x: x.blocks[0].rows[0]._raw["Day"],
		groupWrapper = Day)
	
	return days

# Map[DayNumber -> Map[GroupNumber -> Map[BlockNumber -> Map[RowId -> Row]]]]

days = associateScheduleDayGroupBlock(scheduleWithAbstracts)


printedRoomNames = {
"5th Floor" : "5th Floor",
"Ball-A" : "Ballroom A",
"Ball-B" : "Ballroom B",
"Combined" : "Combined Ballrooms",
"Crystal" : "Crystal Ballroom"
}

#Functions for building HTML for Website and for Microsoft Word import
def renderBlockHeader(block):
	startTime = block.startTime()
	endTime = block.endTime()
	session = block.sessionName() + ": " + block.name()
	room = block.room()
	if room in printedRoomNames:
		printedRoomName = "(" + printedRoomNames[room] + ")"
	else: 
		printedRoomName = ""	
	return r'<p><span style="font-size:16px;"><strong>%s</strong> %s to %s <i>%s</i></span></p>' % (
		session, startTime, endTime, printedRoomName)

def renderRow(row):
	timeDesc = row.startTime()
	timeDesc = timeDesc.replace(" PM","").replace(" AM", "")
	if row.hasTitleAndAuthor():
		(title, author) = row.titleAndAuthor()
	else:
		(title, author) = ("", "")

	columns = [timeDesc, title, author]
	htmlPieces = ["<td>%s</td>" % column for column in columns]
	return "<tr>\n" + "\n".join(htmlPieces) + "</tr>\n"

def renderPosterHtml(poster):
	output = "<tr>"
	output += r'<td colspan="2">%s (%s)</td>' % (poster.title(), poster.number())
	output += "<td>%s</td>" % poster.authors()
	output += "</tr>"
	return output

def renderBlockToHtml(block):
	outHtml = renderBlockHeader(block)
	if block.hasTalks():
		outHtml += r'<table><col width="5%"><col width="60%"><col width="35%">'
		for row in block.rows:
			if row.isTalk():
				outHtml += renderRow(row)
		sessionName = block.name()
		relatedPosters = postersBySession.get(sessionName, [])
		if not relatedPosters == []:
			outHtml += r'<tr><td colspan="3"; valign="bottom"><h3>Associated Posters</h3></td></tr>'
		for poster in relatedPosters:
			outHtml += renderPosterHtml(poster)
		outHtml += "</table>"
	return outHtml


def renderGroupToHtml(group, dayName):
	groupName = group.name()
	if (dayName, groupName) in printedGroupNames:
		printedGroupName = printedGroupNames[(dayName, groupName)]
		outHtml = r'<h3><span style="font-size:18px;"><i>%s</i></span></h3>' % printedGroupName
	else: 
		outHtml = ""	
	for block in group.blocks:
		outHtml += renderBlockToHtml(block)
	return outHtml

def renderDayToHtml(day):
	dayName = day.name()
	outHtml = r'<p>&nbsp;</p><h2><span style="font-size:20px;">%s</span></h2>' % dayName
	for group in day.groups:
		outHtml += renderGroupToHtml(group, dayName)
	return outHtml

#Final Output
htmlParts = []
for day in days:
	htmlParts.append(renderDayToHtml(day))
htmlScheduleBody = "\n".join(htmlParts)

# This outputs a full html document. It is openable in Word
with open("HTMLforWordImport.html","w") as f:
	f.write("<html><body>" + htmlScheduleBody + "</body></html>")

# This outputs just the html for the schedule.
with open("scheduleFragment.html", "w") as f:
	f.write(htmlScheduleBody)
