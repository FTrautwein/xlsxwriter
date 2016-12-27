#define __XAILER_APP__
#ifdef __XAILER_APP__
#include "xailer.ch"
#else
#include "hbclass.ch"
#include "common.ch"
#endif
#include "fileio.ch"
#DEFINE pCRLF CHR(13)+CHR(10)

//------------------------------------------------------------------------------

CLASS XlsxWriter INHERIT XlsxWriterPython

   DATA cLogDir INIT CurrentDir()+"log\"
   DATA cLogFile INIT "python.log"
   DATA lLog INIT .t.
   DATA lTemp INIT .f.

   METHOD New(cWorkBookFile,cPyFile)
   METHOD RunPython()
   METHOD Log()

END CLASS

//------------
METHOD New(cWorkBookFile,cPyFile) CLASS XlsxWriter
LOCAL cPython:= "", nH
IF FILE("c:\python\python.exe")
   ::cPythonExe:= "c:\python\python.exe"
ENDIF
IF EMPTY(cPyFile)
   #ifdef __XAILER__
   ::cPyFile:= FileUnique( HB_DIRTEMP(), "py")
   ::nH:= FCreate( ::cPyFile )
   #else
   ::nH:= HB_FTEMPCREATEEX( @::cPyFile, HB_DIRTEMP(), "XA_", ".py" )
   #endif
ELSE
   ::cPyFile:= cPyFile
   ::nH:= FCreate( ::cPyFile )
ENDIF
IF FERROR() != 0
   RETURN Nil
ENDIF
IF EMPTY(cWorkBookFile)
   ::lTemp:= .t.
   #ifdef __XAILER__
   ::cWorkBookFile:= FileUnique( HB_DIRTEMP(), "xlsx")
   #else
   nH:= HB_FTEMPCREATEEX( @::cWorkBookFile, HB_DIRTEMP(), "XA_", ".xlsx" )
   FCLOSE(nH)
   #endif
ELSE
   ::lTemp:= .f.
   ::cWorkBookFile:= cWorkBookFile
ENDIF
cPython:= ''
cPython+= 'import xlsxwriter'+pCRLF
cPython+= 'from datetime import date'+pCRLF
cPython+= 'Workbook = xlsxwriter.Workbook("'+STRTran(::cWorkBookFile,"\","\\")+'")'+pCRLF
FWrite( ::nH, Hb_StrToUTF8(cPython ) )
RETURN Self

//------------
METHOD RunPython() CLASS XlsxWriter
LOCAL nExitCode, cCmd, cTemp:= "", nH

cCmd:= ::cPythonExe+' "'+::cPyFile+'"'
#ifdef __XAILER__
 nExitCode:= Execute( cCmd, CurrentDir(), .t., SW_HIDE )
 IF ::lLog
   ::Log( cCmd+" "+HB_NTOS(nExitCode),.t.)
 ENDIF
 IF !Empty(nExitCode)
    IF ::lLog
       cTemp:= GetTempFilename( CurrentDir()+::cLogDir, "PY_" )
       COPYFILE( ::cPyFile, cTemp )
       ::Log( cCmd+" "+HB_NTOS(nExitCode)+" Copied to"+cTemp,.t.)
    ENDIF
 ENDIF
#else
 RunShell( cCmd,,,,.t. )
 ::Log(cCmd,.t.)
#endif
RETURN Nil

//------------
METHOD Log(cMsg,lCompact) CLASS XlsxWriter
LOCAL nCount, nH, cFile

DEFAULT lCompact TO .f.

cFile:= ::cLogDir+::cLogFile
IF !FILE( cFile )
   nH:= FCreate( cFile )
   FClose( nH )
ENDIF
nH:= FOPEN( cFile, FO_WRITE, FO_SHARED )
FSeek( nH, 0, FS_END )
FWrite( nH, CHR(13)+CHR(10) )
FWrite( nH, DTOC(DATE()) )
FWrite( nH, " " )
FWrite( nH, TIME() )
FWrite( nH, " ")
FWrite( nH, cMsg )
FWrite( nH, CHR(13)+CHR(10) )
IF !lCompact
   nCount:= 0
   While !Empty( Procname( ++ nCount ) )
      FWrite( nH, "  " + Procname( nCount ) + " (" + Ltrim( Str( Procline( nCount ), 20 ) ) + ")"  )
      FWrite( nH, CHR(13)+CHR(10) )
   ENDDO
ENDIF

FClose( nH )

RETURN(NIL)


