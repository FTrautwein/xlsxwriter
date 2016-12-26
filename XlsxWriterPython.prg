/*
XlsxWriterPython.prg
Fausto Di Creddo Trautwein, ftwein@gmail.com
Python class: https://github.com/jmcnamara/XlsxWriter
*/

#include "hbclass.ch"
#include "common.ch"
#include "fileio.ch"
#DEFINE pCRLF CHR(13)+CHR(10)

//------------------------------------------------------------------------------

CLASS XlsxWriterPython

   DATA nH
   DATA cPythonExe INIT "bin\python\python.exe"
   DATA cWorkBookFile
   DATA cPyFile INIT "xlsxwriter_python.py"
   DATA aFormat INIT { => }
   DATA nIndexFormat INIT 0
   DATA cWorkSheet
   DATA nColumnAdjustment INIT 0
   DATA nLineAdjustment  INIT 0

   METHOD New(cWorkBookFile,cPyFile)
   METHOD Add_Worksheet( cWs, cNome )
   METHOD Select_Worksheet( cWs )
   METHOD SetFormat( aFormatDef, cFormatName )
   METHOD Write( nRow, nCol, xText, cFormat )
   METHOD Merge( nFirst_Row, nFirst_Col, nLast_Row, nLast_Col, xText, cFormat)
   METHOD Set_Column( nInitialCol, nFinalCol, nWidth, cFormat )
   METHOD Set_Row( nRow, nHeight, cFormat )
   METHOD SaveWorkBook()
   METHOD RunPython()
   METHOD ShowWorkBook()

END CLASS

//------------
METHOD New(cWorkBookFile,cPyFile) CLASS XlsxWriterPython
LOCAL cPython:= "", nH
::cPyFile:= cPyFile
::nH:= FCreate( ::cPyFile )
IF FERROR() != 0
   RETURN Nil
ENDIF
::cWorkBookFile:= cWorkBookFile
cPython+= 'import xlsxwriter'+pCRLF
cPython+= 'workbook = xlsxwriter.Workbook("'+STRTran(::cWorkBookFile,"\","\\")+'")'+pCRLF
FWrite( ::nH, Hb_StrToUTF8(cPython ) )
RETURN Self

//------------
METHOD Add_Worksheet( cWs, cNome ) CLASS XlsxWriterPython
LOCAL cPython:= ""
LOCAL cAux:= ""
cPython:= cWs+" = workbook.add_worksheet(<NAME>)"+pCRLF
IF !EMPTY( cNome )
   cAux:= "'"+cNome+"'"
ENDIF
cPython:= STRTRAN( cPython, "<NAME>", cAux )
FWrite( ::nH, Hb_StrToUTF8(cPython ) )
::Select_Worksheet( cWs )
RETURN Self

//------------
METHOD Select_Worksheet( cWs ) CLASS XlsxWriterPython
::cWorkSheet:= cWs
RETURN Self

//------------
METHOD SetFormat( aFormatDef, cFormatName ) CLASS XlsxWriterPython
LOCAL cPython:= "", cText, cTextFormat, xVar, cKey

DEFAULT cFormatName TO ""

cTextFormat:= ""
FOR EACH xVar IN aFormatDef
   IF !EMPTY(cTextFormat)
      cTextFormat+=", "
   ENDIF
   IF xVar:__enumKey == "valign"
      cKey:= "align"
   ELSE
      cKey:= xVar:__enumKey
   ENDIF
   cTextFormat+= '"'+cKey+'": ' + xVar:__enumValue
NEXT

IF EMPTY(cFormatName ) .AND. HHasKey(::aFormat,cTextFormat)
   cFormatName:= ::aFormat[ cTextFormat ]
ELSE
   //'headerformat = workbook.add_format("bold":True, "bg_color": "#A5A5A5", "text_wrap": True, "border": 1)'+pCRLF
   IF EMPTY(cFormatName)
      cFormatName:= "formato"+HB_NTOS(++::nIndexFormat)
   ENDIF
   ::aFormat[cTextFormat]:= cFormatName
   cPython+= cFormatName+' = workbook.add_format( { '+cTextFormat+' } )'+pCRLF
   FWrite( ::nH, Hb_StrToUTF8(cPython ) )
ENDIF

RETURN cFormatName

//------------
METHOD Write( nRow, nCol, xText, cFormat ) CLASS XlsxWriterPython
LOCAL cPython:= "", xValue:= ""
xText:= IIF( xText == Nil, "", xText )
IF VALTYPE(xText) == "N"
   xValue:= HB_NTOS(xText)
ELSEIF VALTYPE(xText) == "C"
   xValue:= "'"+STRTRAN(xText,"'","\'")+"'"
   xValue:= STRTRAN(xValue,pCRLF,"\n")
ELSEIF VALTYPE(xText) == "D"
   xValue:= "'"+DTOC(xText)+"'"
ELSEIF VALTYPE(xText) == "L"
   xValue:= "'"+IIF(xText,"S","N")+"'"
ELSE
   xValue:= "''"
ENDIF

cPython+= ::cWorkSheet+'.write( '
cPython+= HB_NTOS(nRow+::nLineAdjustment)
cPython+= ',' + HB_NTOS(nCol+::nColumnAdjustment)
cPython+= ', ' + xValue
IF !EMPTY( cFormat )
   cPython+= ', ' + cFormat
ENDIF
cPython+= " )"+pCRLF
FWrite( ::nH, Hb_StrToUTF8( cPython ) )
RETURN Self

//------------
METHOD Merge( nFirst_Row, nFirst_Col, nLast_Row, nLast_Col, xText, cFormat) CLASS XlsxWriterPython
//------------
LOCAL cPython:= "", xValue:= ""
xText:= IIF( xText == Nil, "", xText )
IF VALTYPE(xText) == "N"
   xValue:= HB_NTOS(xText)
ELSEIF VALTYPE(xText) == "C"
   xValue:= "'"+STRTRAN(xText,"'","\'")+"'"
ELSEIF VALTYPE(xText) == "D"
   xValue:= "'"+DTOC(xText)+"'"
ELSEIF VALTYPE(xText) == "L"
   xValue:= "'"+IIF(xText,"S","N")+"'"
ELSE
   xValue:= "''"
ENDIF

cPython+= ::cWorkSheet+'.merge_range( '
cPython+= HB_NTOS(nFirst_Row+::nLineAdjustment)
cPython+= ', ' + HB_NTOS(nFirst_Col+::nColumnAdjustment)
cPython+= ', ' + HB_NTOS(nLast_Row+::nLineAdjustment)
cPython+= ', ' + HB_NTOS(nLast_Col+::nColumnAdjustment)
cPython+= ', ' + xValue
IF !EMPTY( cFormat )
   cPython+= ', ' + cFormat
ENDIF
cPython+= " )"+pCRLF
FWrite( ::nH, Hb_StrToUTF8( cPython ) )
RETURN Self

//------------
METHOD Set_Column(nInitialCol, nFinalCol, nWidth, cFormat ) CLASS XlsxWriterPython
LOCAL cPython:= "", cAux
//ws1.set_column( <COL>, <COL>, <LARGURA> )'+pCRLF
cPython+= ::cWorkSheet+'.set_column( '
cPython+= HB_NTOS(nInitialCol+::nColumnAdjustment)
cPython+= ',' + HB_NTOS(nFinalCol+::nColumnAdjustment)
cPython+= ',' + HB_NTOS(nWidth)
IF !EMPTY( cFormat )
   cPython+= ', ' + cFormat
ENDIF
cPython+= " )"+pCRLF
FWrite( ::nH, Hb_StrToUTF8( cPython ) )
RETURN Self

//------------
METHOD Set_Row( nRow, nHeight, cFormat ) CLASS XlsxWriterPython
LOCAL cPython:= "", cAux
//set_row(row, height, cell_format, options)
cPython+= ::cWorkSheet+'.set_row( '
cPython+= HB_NTOS(nRow+::nLineAdjustment)
cPython+= ',' + HB_NTOS(nHeight)
IF !EMPTY( cFormat )
   cPython+= ', ' + cFormat
ENDIF
cPython+= " )"+pCRLF
FWrite( ::nH, Hb_StrToUTF8( cPython ) )
RETURN Self

//------------
METHOD SaveWorkBook() CLASS XlsxWriterPython
LOCAL cPython:= ""
cPython+= 'workbook.close()'
FWrite( ::nH, Hb_StrToUTF8( cPython ) )
FClose( ::nH )
RETURN Nil

//------------
METHOD RunPython() CLASS XlsxWriterPython
LOCAL nExitCode, cCmd, cTemp:= "", nH

cCmd:= ::cPythonExe+' "'+::cPyFile+'"'
RunShell( cCmd,,,,.t. )
RETURN Nil

//------------
METHOD ShowWorkBook() CLASS XlsxWriterPython
LOCAL nExitCode
nExitCode:= WAPI_ShellExecute(::cWorkBookFile,, ::cWorkBookFile,,, 1 )
RETURN Nil


