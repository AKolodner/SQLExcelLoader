import sys
import pickle
import sqlite3
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
    sys.exit()
elif sys.argv[1] == "download": # --DOWNLOAD--
    # have option to download current active sheet or enter a name
    if len(sys.argv) < 4:
        print("Usage: python3 file.py download * <filePath>\nOr: python3 file.py download <sheetName> <filePath>")
        sys.exit()
    try:
        app = xlw.apps[xlw.apps.keys()[0]]
    except IndexError:
        app = xlw.App()

    try:
        if sys.argv[2] == "-":
            spread = app.books.active
            print("Using active sheet " + spread.name + "...")
        elif sys.argv[2] != "":
            spread = app.books[sys.argv[2]]
        else:
            print("Enter a sheet name (or * to use the active sheet)!")
            sys.exit(0)
    except:
        print("ERROR: Sheet not found!\nUsage: python3 file.py download * <filePath> \nOr: python3 file.py download <sheetName> <filePath>")
        sys.exit()

    try:
        file = open(sys.argv[3], 'w')
    except FileNotFoundError:
        print("ERROR: Directory to save file to could not be found.")
        sys.exit()

    file.write("BEGIN TRANSACTION;\n")

    tables = []
    sheetMetadata = {}

    try:
        sheetMetadata = pickle.loads(bytes.fromhex(spread.sheets['__METADATA'].range('E1').value))
    except:
        print("No metadata found or metadata corrupted.")
        sys.exit()
    
    for table in spread.sheets:
        if table.name == "__METADATA":
            continue
    
        tables.append(table.name)


        file.write('CREATE TABLE \"' + table.name + '" (\n')

        primaryKeyItems = []

        first = True

        for field in sheetMetadata[table.name]:
            if first:
                first = False
            else:
                file.write(",\n")
            
            if field[5] > 0:
                primaryKeyItems.append(field[1])
            file.write('\t"' + field[1] + '" ' + field[2])
            if field[3]:
                file.write(' NOT NULL')

        file.write(",\n\tPRIMARY KEY (")
        first = True
        for item in primaryKeyItems:
            if first:
                first = False
            else:
                file.write(", ")
            file.write(item)
        file.write(")\n);\n")

        usedCells = table.range('A1').expand().value[1:]

        for row in usedCells:
            file.write('INSERT INTO "' + table.name + '" VALUES(')

            first = True
            for cell in row:
                if first:
                    first = False
                else:
                    file.write(',')

                if cell == 'NULL':
                    file.write('NULL')
                elif cell == "false":
                    file.write("False")
                elif cell is None:
                    file.write("''")
                else:
                    file.write("'" + str(cell).replace("'","''") + "'")

            file.write(');\n')
            
    file.write("COMMIT;")
    file.close()
        
elif path.isfile(sys.argv[1]): #--UPLOAD--
    filename = sys.argv[1].split("/")[len(sys.argv[1].split("/")) - 1]

    sqlfile = open(sys.argv[1], 'r')

    conn = sqlite3.connect(':memory:')
    cursor = conn.cursor()

    cursor.executescript(sqlfile.read())

    # This command queries SQLite for the names of all tables that were just generated
    cursor.execute("select name from sqlite_master where type = 'table';")
    # The query returns a list of one-item tuples, so this creates a list of just the contents (for ease of use)
    tables = [entry[0] for entry in cursor.fetchall()]

    # sheetMetadata will hold information on the type of each field in each table, so it can be saved back from Excel.
    sheetMetadata = {}

    try:
        app = xlw.apps[xlw.apps.keys()[0]]
    except IndexError:
        app = xlw.App()
        
    spread = app.books.add()
    sheetMetadata = {}

    for table in tables:
        cursor.execute("select * from '" + table + "'")
        # Creates a list of field names (for writing to Excel)
        fields = [item[0] for item in cursor.description]

        currSheet = spread.sheets.add(table, after=spread.sheets[len(spread.sheets)-1])
        currSheet.range('A1').value = fields

        cursor.execute("select * from '" + table + "'")
        tableData = cursor.fetchall()

        row = 2
        for entry in tableData:
            currSheet.range('A' + str(row)).value = entry

            column = 1
            for cell in entry:
                if cell is None:
                    currSheet.range(columnLetter(column) + str(row)).value = 'NULL'
                column += 1
                    
            
            row += 1

        cursor.execute("pragma table_info('" + table + "')")
        sheetMetadata[table] = cursor.fetchall()

    currSheet = spread.sheets.add("__METADATA", after=spread.sheets[len(spread.sheets)-1])
    currSheet.range('A1','Z1').color = (255,0,0)
    currSheet.range('A1','A5').value = ('DO NOT','EDIT','THIS','SHEET',pickle.dumps(sheetMetadata).hex())

    try:
        spread.sheets["Sheet1"].delete()
    except:
        pass
    spread.sheets[0].activate()
