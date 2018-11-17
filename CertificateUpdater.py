from openpyxl import Workbook
from openpyxl import load_workbook

def printSheetRow (row, rowType, fileName):
    """"Print a row of a spreadsheet"""
    print ("\n", rowType, " from", fileName, ":")
    for cell in row:
        print (cell, " = ", cell.value)
    return 1

def findCellInRow (row, cellValue):
    """Seach a row to find a value and return the column index."""
    for cell in row:
        if cell.value == cellValue:
            return cell.col_idx
    return None
    
def findCellRowInColumn (targetSheet, column, cellValue):
    """Seach a column to find a value and return the row index."""
    for row in range (1,targetSheet.max_row+1):
        if targetSheet.cell(row=row, column=column).value == cellValue:
            return targetSheet.cell(row=row, column=column).row
    return None

def quickcheck(a):
    """ """
    a.cell(1,1,value=3)
    return None

# path to files
path = "c:/Users/gja/Documents/py_test/"

# Header file
headerFile = "headers.xlsx"

# kickout file
kickoutFile = "Kickouts.xlsx"

# cvent contacts
cventContactFile = "CventContacts.xlsx"

# load the headers
header = load_workbook(path+headerFile)
headerSheet = header['Sheet1']
headerHeaderRow = headerSheet[1]
printSheetRow (headerHeaderRow,"Headers", headerFile)
print (findCellInRow(headerHeaderRow, "Seven"))
print ("max row = ",headerSheet.max_row)
print ("max col = ",headerSheet.max_column)
quickcheck(headerSheet)
print (headerSheet.cell(1,1), " = ",headerSheet.cell(1,1).value)

# load the kickout fileheader
kickout = load_workbook(path+kickoutFile)
kickoutSheet = kickout['Sheet1']
kickoutHeaderRow = kickoutSheet[1]
printSheetRow (kickoutHeaderRow,"Headers", kickoutFile)

# load the cvent contact fileheader 
cventContact = load_workbook(path+cventContactFile)
cventContactSheet = cventContact['Sheet1']
cventContactHeaderRow = cventContactSheet[1]
printSheetRow (cventContactHeaderRow,"Headers", cventContactFile)

# create a new worksheet to map the required column from kickout to the contact colums
# row 1 will be the headers from the kickout file
# row 2 will be the corresponding column index of the contact file
columnWB = Workbook()
columnMap = columnWB.active
for cell in kickoutHeaderRow:
    columnMap.cell(row=1,column=cell.col_idx,value=cell.value)
    columnMap.cell(row=2,column=cell.col_idx,value=findCellInRow(cventContactHeaderRow, cell.value))
printSheetRow (columnMap[1],"Headers", "columnMap")
printSheetRow (columnMap[2],"Contact Columns", "columnMap")

# Set the column for the matching
kickoutKeyColumn = findCellInRow(kickoutHeaderRow, "One")
contactKeyColumn = findCellInRow(cventContactHeaderRow, "One")

# Find the kickout values in the contacts
# once found build the using contact values to override kickout values
col = kickoutKeyColumn
for row in range (2,kickoutSheet.max_row+1):
    #print("row=",row," column=",col)
    kickoutCell = kickoutSheet.cell(row=row,column=col)
    rowMatch = findCellRowInColumn(cventContactSheet,contactKeyColumn,kickoutCell.value)
    # start a new row
    newRow=columnMap.max_row+1
    for setCol in range (1,columnMap.max_column+1):
        columnMap.cell(row=newRow, column=setCol, value=kickoutSheet.cell(row=row,column=setCol).value)
        if rowMatch is not None:
            cventCell=cventContactSheet.cell(row=rowMatch,column=contactKeyColumn)
            if cventCell.value is not None:
                columnMap.cell(row=newRow, column=setCol, value=cventCell.value)
                
print ("max row = ",columnMap.max_row)
print ("max col = ",columnMap.max_column)

print ("Done")




