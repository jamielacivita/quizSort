from openpyxl import load_workbook

DEBUG = False
if __name__ == "__main__":
	DEBUG = False

#load roster data
def load_roster_data():
	return "load roster data"


#load quiz data
def load_quiz_data():
	#import to read xlsx in python.
	#get the 
	#Open PQ.xlsx
	#Switch to Sheet 'Player Level'
	#Read in cell A5
	return "load quiz data"


#generate data student / score / class
def load_data_student_score_class():
	return "load data student score class"
	#if student does not have a score put zero.


#generate csv file(s)
def generate_csv_files():
	return "generate csv files"


#TODO: this is a test funciton - clean it up
def readCellA5():
	#import to read xlsx in python.
	wb = load_workbook(filename = r"c:\source\python\quizSort\Data\PQ.xlsx")
	if DEBUG : print("WB: %s" % wb)
	if DEBUG : print("Sheet Names: %s" % wb.sheetnames)
	#OQ: how to switch to Sheet 'Player Level'
	#ws = wb.get_sheet_by_name("Player Level")
	ws = wb["Player Level"]
	if DEBUG : print("Worksheet : %s " % ws)

	#OQ: How to tead in cell A5

	result = ws['A5'].value
	if DEBUG : print("Result: %s " % result)
	return result


def readCellB2():
	#import to read xlsx in python.
	wb = load_workbook(filename = r"c:\source\python\quizSort\Data\PQ.xlsx")
	if DEBUG : print("In read cell B2 for sheet 1st")
	if DEBUG : print("WB: %s" % wb)
	if DEBUG : print("Sheet Names: %s" % wb.sheetnames)
	#OQ: how to switch to Sheet 'Player Level'
	#ws = wb.get_sheet_by_name("1st")
	ws = wb["1st"]
	if DEBUG : print("Worksheet : %s " % ws)

	#OQ: How to tead in cell A5

	result = ws['B2'].value
	result = result.upper()
	if DEBUG : print("Result: %s " % result)
	return result


def getFirstClassStudents_lst():
	wb = load_workbook(filename = r"c:\source\python\quizSort\Data\PQ.xlsx")
	ws = wb["1st"]
	#Students in the first period class go from B2 to B25
	students_tpl = ws["B2":"B25"]
	students_lst = list(students_tpl)
	students_lst[:] = [s[0].value.upper() for s in students_lst]
	if DEBUG : print("students_lst: %s " % students_lst)
	return students_lst


def writeTable():
	fileName = r"c:\source\python\quizSort\Data\class07.txt"
	print("writing table file.")
	students_lst = getFirstClassStudents_lst()
	with open(fileName, "w") as f:
		for s in students_lst:
			f.write(s + "\n")

def parsePlayers(raw_player):
	if DEBUG : print("In parsePlayers")
	player = raw_player.split("(")[1]
	if DEBUG : print("player:  %s" % player)
	player = player.split(")")[0]
	player = player.strip()
	if DEBUG : print("player: %s" % player)
	return player.upper()

	
def parsePlayerLevelR5():
	#import to read xlsx in python.
	wb = load_workbook(filename = r"c:\source\python\quizSort\Data\PQ.xlsx")
	if DEBUG : print("WB: %s" % wb)
	if DEBUG : print("Sheet Names: %s" % wb.sheetnames)
	ws = wb["Player Level"]
	if DEBUG : print("Worksheet : %s " % ws)
	raw_players = ws['A5'].value
	player = parsePlayers(raw_players)
	raw_score = ws['B5'].value
	score = str(raw_score)
	raw_accuracy = ws['C5'].value
	accuracy = str(int(float(raw_accuracy)*100))
	accuracy = accuracy + "%"
	if DEBUG : print("player: %s " % player)
	if DEBUG : print("score: %s " % score)
	if DEBUG : print("accuracy: %s " % accuracy)
	return (player, score, accuracy)

	

def getParsePlayerLevelRow(row_number):
	#import to read xlsx in python.
	wb = load_workbook(filename = r"c:\source\python\quizSort\Data\PQ.xlsx")
	if DEBUG : print("WB: %s" % wb)
	if DEBUG : print("Sheet Names: %s" % wb.sheetnames)
	ws = wb["Player Level"]
	if DEBUG : print("Worksheet : %s " % ws)
	raw_players = ws.cell(row = row_number, column=1).value
	player = parsePlayers(raw_players)
	raw_score = ws.cell(row = row_number, column=2).value
	score = str(raw_score)
	raw_accuracy = ws.cell(row = row_number, column = 3).value
	accuracy = str(int(float(raw_accuracy)*100))
	accuracy = accuracy + "%"
	if DEBUG : print("player: %s " % player)
	if DEBUG : print("score: %s " % score)
	if DEBUG : print("accuracy: %s " % accuracy)
	return (player, score, accuracy)


def parsePlayer():
	START_ROW = 5
	END_ROW = 90
	player_level_dict = {}
	for row in range(START_ROW, END_ROW+1):
		key = getParsePlayerLevelRow(row)[0]
		value = getParsePlayerLevelRow(row)
		player_level_dict[key]=value
	return player_level_dict








def main():
	#print("JWTO")
	#readCellA5()
	#getFirstClassStudents_lst()
	#writeTable()
	#parsePlayerLevelR5()
	#parsePlayerLevel(5)
	
	print(parsePlayer())
	#parseFirstPeriod()




if __name__ == "__main__":
	main()
