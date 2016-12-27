FUNCTION Main()
LOCAL aExpenses, row, col, oXlsx, i

// Some data we want to write to the worksheet.
aExpenses:= {;
    {'Rent', 1000},;
    {'Gas',   100},;
    {'Food',  300},;
    {'Gym',    50};
}

// Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

oXlsx:= XlsxWriterPython():New("example1.xlsx","example1.py")
//oXlsx:cPythonExe:= "c:\python\python.exe"
oXlsx:Add_Worksheet( "ws1", "sheet1" )

// Iterate over the data and write it out row by row.
FOR i:= 1 TO LEN( aExpenses )
   oXlsx:Write( row, col    , aExpenses[i,1] )
   oXlsx:Write( row, col + 1, aExpenses[i,2] )
   row += 1
NEXT

// Write a total using a formula.
oXlsx:Write( row, 0, 'Total' )
oXlsx:Write( row, 1, '=SUM(B1:B4)' )

oXlsx:SaveWorkBook()
oXlsx:RunPython()
oXlsx:ShowWorkBook()

RETURN Nil
