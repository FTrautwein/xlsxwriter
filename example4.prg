FUNCTION Main(nLines)
LOCAL row, oXlsx, i, cName, cDate, cTitle, aDoc, nL, nC

nLines:= IIF( nLines == Nil, 40, VAL(nLines) )

Set( _SET_DATEFORMAT, "yyyy-mm-dd" )

// Start from the first cell. Rows and columns are zero indexed.
row = 0

oXlsx:= XlsxWriterPython():New( "example4.xlsx", "example4.py" )
//oXlsx:cPythonExe:= "c:\python\python.exe"
oXlsx:Add_Worksheet( "ws1", "sheet1" )

SetFormat( oXlsx )

//Adjust the column width.
oXlsx:Set_Column(  0,  0,  14 ) // Invoice
oXlsx:Set_Column(  1,  1,   4 ) // TP
oXlsx:Set_Column(  2,  2,  14 ) // Date Mov.
oXlsx:Set_Column(  3,  3,  14 ) // Date Issue
oXlsx:Set_Column(  4,  4,  10 ) // Tax Code
oXlsx:Set_Column(  5,  5,  10 ) // Cust.Code
oXlsx:Set_Column(  6,  6,  40 ) // Customer
oXlsx:Set_Column(  7,  7,   4 ) // ST
oXlsx:Set_Column(  8,  8,  16 ) // Total Value
oXlsx:Set_Column(  9,  9,  16 ) // Base Value1
oXlsx:Set_Column( 10, 10,  16 ) // Tax/Fee1
oXlsx:Set_Column( 11, 11,  16 ) // Base Value2
oXlsx:Set_Column( 12, 12,  16 ) // Tax/Fee2

cName := "ENTERPRISE NAME"
cDate := "2016-12-26"
cTitle:= "HARBOUR XLSXWRITER PYTHON REPORT"

oXlsx:Merge(   row, 0, row, 12, cName, 'Header' )
oXlsx:Merge( ++row, 0, row, 12, "Date:"+cDate,'Header')
oXlsx:Merge( ++row, 0, row, 12, cTitle,'Header')

oXlsx:Write( ++row,  0, "Invoice"           , 'textLeftBoldColor' )
oXlsx:Write(   row,  1, "TP"                , 'textLeftBoldColor' )
oXlsx:Write(   row,  2, "Date Mov."         , 'textLeftBoldColor' )
oXlsx:Write(   row,  3, "Date Issue"        , 'textLeftBoldColor' )
oXlsx:Write(   row,  4, "Tax Code"          , 'textLeftBoldColor' )
oXlsx:Write(   row,  5, "Cust.Code"         , 'textLeftBoldColor' )
oXlsx:Write(   row,  6, "Customer"          , 'textLeftBoldColor' )
oXlsx:Write(   row,  7, "ST"                , 'textLeftBoldColor' )
oXlsx:Write(   row,  8, "Total Value"       , 'textRightBoldColor' )
oXlsx:Write(   row,  9, "Base Value1"       , 'textRightBoldColor' )
oXlsx:Write(   row, 10, "Tax/Fee1"          , 'textRightBoldColor' )
oXlsx:Write(   row, 11, "Base Value2"       , 'textRightBoldColor' )
oXlsx:Write(   row, 12, "Tax/Fee2"          , 'textRightBoldColor' )

aDoc := GetData(nLines)

QOUT()
nL:= ROW()
nC:= COL()

FOR i := 1 TO nLines
   @ nL,nC SAY i/nLines*100
   oXlsx:Write( ++row,  0, aDoc[ i, 1 ], "textLeft" )
   oXlsx:Write(   row,  1, aDoc[ i, 2 ], "textLeft" )
   oXlsx:Write(   row,  2, DToC( aDoc[ i, 3 ] ), "textLeft" )
   oXlsx:Write(   row,  3, DToC( aDoc[ i, 4 ] ), "textLeft" )
   oXlsx:Write(   row,  4, aDoc[ i, 5 ], "textLeft" )
   oXlsx:Write(   row,  5, aDoc[ i, 6 ], "textLeft" )
   oXlsx:Write(   row,  6, aDoc[ i, 7 ], "textLeft" )
   oXlsx:Write(   row,  7, aDoc[ i, 8 ], "textLeft" )
   oXlsx:Write(   row,  8, aDoc[ i, 9 ], "numberRight" )
   oXlsx:Write(   row,  9, aDoc[ i, 10 ], "numberRight" )
   oXlsx:Write(   row, 10, aDoc[ i, 11 ], "numberRight" )
   oXlsx:Write(   row, 11, aDoc[ i, 12 ], "numberRight" )
   oXlsx:Write(   row, 12, aDoc[ i, 13 ], "numberRight" )
NEXT

oXlsx:Write( ++row,  0, "", "textLeft" )
oXlsx:Write(   row,  1, "", "textLeft" )
oXlsx:Write(   row,  2, "", "textLeft" )
oXlsx:Write(   row,  3, "", "textLeft" )
oXlsx:Write(   row,  4, "", "textLeft" )
oXlsx:Write(   row,  5, "", "textLeft" )
oXlsx:Write(   row,  6, "TOTAL ==> " + hb_ntos( LEN(aDoc) ) + " document(s)", "textLeftBold" )
oXlsx:Write(   row,  7, "", "textLeft" )
oXlsx:Write(   row,  8, "=SUM(I5:I"+HB_NTOS(nLines+4)+")", "numberRightBold" )
oXlsx:Write(   row,  9, "=SUM(J5:J"+HB_NTOS(nLines+4)+")", "numberRightBold" )
oXlsx:Write(   row, 10, "=SUM(K5:K"+HB_NTOS(nLines+4)+")", "numberRightBold" )
oXlsx:Write(   row, 11, "=SUM(L5:L"+HB_NTOS(nLines+4)+")", "numberRightBold" )
oXlsx:Write(   row, 12, "=SUM(M5:M"+HB_NTOS(nLines+4)+")", "numberRightBold" )

oXlsx:Add_Worksheet( "ws2", "sheet2" )
oXlsx:Write( 0,  0, "TEXT LIN 1 COL 1 SHEET2", "textLeft" )
oXlsx:Write( 1,  0, "TEXT LIN 2 COL 1 SHEET2", "textLeft" )

oXlsx:SaveWorkbook()
oXlsx:RunPython()
oXlsx:ShowWorkbook()

RETURN Nil

FUNCTION SetFormat( oXlsx )
LOCAL aFormat

oXlsx:SetFormat( { 'align' => '"left"', 'valign' => '"vcenter"', 'font_size' => '11' }, 'textLeft' )
oXlsx:SetFormat( { 'align' => '"left"', 'valign' => '"vcenter"', 'font_size' => '11', 'bold' => 'True' }, 'textLeftBold' )
oXlsx:SetFormat( { 'align' => '"left"', 'valign' => '"vcenter"', 'font_size' => '11', 'bold' => 'True', 'bg_color' => '"#A5A5A5"' }, 'textLeftBoldColor' )

oXlsx:SetFormat( { 'align' => '"right"', 'valign' => '"vcenter"', 'font_size' => '11' }, 'textRight' )
oXlsx:SetFormat( { 'align' => '"right"', 'valign' => '"vcenter"', 'font_size' => '11', 'bold' => 'True' }, 'textRightBold' )

aFormat:= { => }
aFormat[ 'font_size' ]:= '11'
aFormat[ 'bold'      ]:= 'True'
aFormat[ 'valign'     ]:= '"vcenter"'
aFormat[ 'align'     ]:= '"right"'
aFormat[ 'bg_color'  ]:= '"#A5A5A5"'
aFormat[ 'text_wrap' ]:= 'True'
oXlsx:SetFormat( aFormat, 'textRightBoldColor' )

aFormat:= { => }
aFormat[ 'font_size'  ]:= '11'
aFormat[ 'valign'     ]:= '"vcenter"'
aFormat[ 'align'      ]:= '"right"'
aFormat[ 'num_format' ]:= '"#,##0.00"' 
oXlsx:SetFormat( aFormat, 'numberRight' )

aFormat:= { => }
aFormat[ 'font_size'  ]:= '11'
aFormat[ 'valign'     ]:= '"vcenter"'
aFormat[ 'align'      ]:= '"right"'
aFormat[ 'num_format' ]:= '"#,##0.00"' 
aFormat[ 'bold'       ]:= 'True'
oXlsx:SetFormat( aFormat, 'numberRightBold' )

aFormat:= { => }
aFormat[ 'font_size'  ]:= '11'
aFormat[ 'valign'     ]:= '"vcenter"'
aFormat[ 'align'      ]:= '"right"'
aFormat[ 'num_format' ]:= '"#,##0.00"' 
aFormat[ 'bold'       ]:= 'True'
aFormat[ 'bg_color'   ]:= '"#A5A5A5"'
oXlsx:SetFormat( aFormat, 'numberRightBoldColor' )

aFormat:= { => }
aFormat[ 'font_size'  ]:= '13'
aFormat[ 'bold'       ]:= 'True'
aFormat[ 'valign'     ]:= '"vcenter"'
oXlsx:SetFormat( aFormat, 'Header' )

aFormat:= { => }
aFormat[ 'font_size'  ]:= '13'
aFormat[ 'bold'       ]:= 'True'
aFormat[ 'valign'     ]:= '"vcenter"'
oXlsx:SetFormat( aFormat, 'HeaderRight' )

aFormat:= { => }
aFormat[ 'font_size'  ]:= '13'
aFormat[ 'bold'       ]:= 'True'
aFormat[ 'align'      ]:= '"center"'
aFormat[ 'valign'     ]:= '"vcenter"'
oXlsx:SetFormat( aFormat, 'HeaderCenter' )

RETURN Nil

FUNCTION GetData(nLines)
LOCAL aDoc:= {}, i, n
FOR i:= 1 TO nLines
   AAdd( aDoc, ;
      { StrZero( i, 8 ), ;
      "SA", ;
      Date() - 49 - i, ;
      Date() - 50 - i, ;
      "5.102", ;
      StrZero( i, 5 ), ;
      "TEST CUSTOMER NAME " + hb_ntos( i ), ;
      "ME", ;
      n:= HB_Random(10) * 1000, ;
      n * 0.90, ;
      n * 0.90 * 0.12, ;
      n * 0.80, ;
      n * 0.10 } )
NEXT
RETURN aDoc