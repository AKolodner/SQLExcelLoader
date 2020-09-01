import sys
import pickle
import xlwings as xlw
from os import path

def columnLetter(n: int) -> str:
    "Converts a number to its letter equivalent.\nEx: 1->A, 5->E, 29->AC, etc."
    # code source: https://stackoverflow.com/questions/23861680/convert-spreadsheet-number-to-column-letter
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

if len(sys.argv) <= 1:
    print("Please enter a command or file name!")
elif sys.argv[1] == "download":
    # have option to download current active sheet or enter a name
    pass
elif path.isfile(sys.argv[1]):
    filename = sys.argv[1].split("/")[len(sys.argv[1].split("/")) - 1]

    try:
        app = xlw.apps[xlw.apps.keys()[0]]
    except IndexError:
        app = xlw.App()
        
    spread = app.books.add()
    #spread.name = filename

    sheetMetadata = {} # dictionary: each key is a sheet name, each value is a dictionary of fields and their SQL definition lines to make sure it saves correctly

    sqlfile = open(sys.argv[1], 'r')
    
    line = sqlfile.readline()
    while line:
        if line[:12] == "CREATE TABLE":
            sheetName = line[14:len(line)-4]
            currSheet = spread.sheets.add(sheetName, after=spread.sheets[len(spread.sheets)-1])
            sheetMetadata[sheetName] = {}
            
            # iterate over next few lines to add in fields
            line = sqlfile.readline()
            numFields = 0
            while line.strip()[:11] != "PRIMARY KEY":
                fieldLine = line.strip()
                fieldName = fieldLine.split(" ")[0].strip("\"")

                numFields += 1
                currSheet.range(columnLetter(numFields) + '1').value = fieldName
                
                sheetMetadata[sheetName][fieldName] = fieldLine
                
                line = sqlfile.readline()
            sheetMetadata[sheetName]["__PRIMARYKEY"] = line
                
        elif line[:11] == "INSERT INTO":
            print(line)
            returnSheet = currSheet
            if line.split(" ")[2].strip(1) != currSheet.name: # if the insert is to a different table (this shouldn't ever be the case but it can't hurt to account for it)
                currSheet = spread.sheets[line.split(" ")[2].strip("\"")]

            scanMode = "waitforstart"
            scan = ""
            moreToScan = True
            quoteKind = ""
            while moreToScan:
                for char in line:
                    if scanMode == "waitforstart" and char == "(":
                        scanMode = "values"
                    elif scanMode == "values" and char in ("'","\""):
                        scanMode = "reading"
                        quoteKind = char
                    elif scanMode == "reading":
                        pass
                
##            ###########
##            # THIS DOESN'T WORK PROPERLY AND NEEDS TO BE FIXED: can't handle dates or multiline
##            # may have to start reading by character instead of split
##            values = line.split(" ")[3] # values = the VALUES() command in that line
##            values = values[7:len(values)-2] # chops off the preceding "VALUES(" and trailing ");"
##            values = values.split(",") # chops into list
##
##            for field in values:
##                if field[:1] == "'" or field[:1] == "\"":
##                    field = field[1:len(field)-1]
##            ############

            currSheet = returnSheet
        
        line = sqlfile.readline()

    metadataSheet = spread.sheets.add("__METADATA", after=spread.sheets[len(spread.sheets)-1])
    metadataSheet.range('A1').value = [["DO NOT"],["EDIT"],["THIS"],["SHEET"]]
    metadataSheet.range('A5').number_format = "@"
    metadataSheet.range('A5').formula = pickle.dumps(sheetMetadata)
    metadataSheet.range('A1','A5').color = (255,0,0)

    try:
        spread.sheets["Sheet1"].delete()
    except:
        pass
    spread.sheets[0].activate()
