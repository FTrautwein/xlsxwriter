FUNCTION Main()
LOCAL aExpenses, row, col, oXlsx, i

// Some data we want to write to the worksheet.
aExpenses:= {;
    { 'Rent', STOD('20130113'), 1000 },;
    { 'Gas',  STOD('20130114'),  100 },;
    { 'Food', STOD('20130116'),  300 },;
    { 'Gym',  STOD('20130120'),   50 };
}

// Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

oXlsx:= XlsxWriterPython():New( "example3.xlsx", "example3.py" )
//oXlsx:cPythonExe:= "c:\python\python.exe"
oXlsx:Add_Worksheet( "ws1", "sheet1" )

// Add a bold format to use to highlight cells.
oXlsx:SetFormat( { 'bold' => 'True' }, 'bold' )
// Add a bold format to use to highlight cells.
oXlsx:SetFormat( { 'bold' => 'True', 'align' => '"center"' }, 'boldcenter' )
oXlsx:SetFormat( { 'bold' => 'True', 'align' => '"right"' }, 'boldright' )

//Add a number format for cells with money.
oXlsx:SetFormat( {'num_format' => '"$#,##0"'}, 'money_format' )
oXlsx:SetFormat( {'num_format' => '"$#,##0"', 'bold' => 'True'}, 'money_formatbold' )

//Add an Excel date format.
oXlsx:SetFormat( {'num_format' => '"d mmmm yyyy"'}, 'date_format' )

//Adjust the column width.
oXlsx:Set_Column(1, 1, 15)

// Write some data headers.
oXlsx:Write( row, col  , 'Item', 'bold' ) 
oXlsx:Write( row, col+1, 'Date', 'boldcenter' )
oXlsx:Write( row, col+2, 'Cost', 'boldright' )

// Iterate over the data and write it out row by row.
FOR i:= 1 TO LEN( aExpenses )
   oXlsx:Write( ++row, col    , aExpenses[i,1] )
   oXlsx:Write(   row, col + 1, aExpenses[i,2], 'date_format' )
   oXlsx:Write(   row, col + 2, aExpenses[i,3], 'money_format' )
NEXT

// Write a total using a formula.
oXlsx:Write( ++row, 0, 'Total', 'bold' )
oXlsx:Write(   row, 2, '=SUM(C1:C4)', 'money_formatbold' )

oXlsx:SaveWorkBook()
oXlsx:RunPython()
oXlsx:ShowWorkBook()

RETURN Nil
