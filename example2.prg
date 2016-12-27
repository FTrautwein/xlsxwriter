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

oXlsx:= XlsxWriterPython():New("example2.xlsx","example2.py")
//oXlsx:cPythonExe:= "c:\python\python.exe"
oXlsx:Add_Worksheet( "ws1", "sheet1" )

// Add a bold format to use to highlight cells.
oXlsx:SetFormat( { 'bold' => 'True' }, 'bold' )

//Add a number format for cells with money.
oXlsx:SetFormat( {'num_format' => '"$#,##0"'}, 'money' )

// Write some data headers.
oXlsx:Write( row, col  , 'Item', 'bold' ) 
oXlsx:Write( row, col+1, 'Cost', 'bold' )

// Iterate over the data and write it out row by row.
FOR i:= 1 TO LEN( aExpenses )
   oXlsx:Write( ++row, col    , aExpenses[i,1] )
   oXlsx:Write(   row, col + 1, aExpenses[i,2], 'money' )
NEXT

// Write a total using a formula.
oXlsx:Write( ++row, 0, 'Total', 'bold' )
oXlsx:Write(   row, 1, '=SUM(B1:B4)', 'money' )

oXlsx:SaveWorkBook()
oXlsx:RunPython()
oXlsx:ShowWorkBook()

RETURN Nil
