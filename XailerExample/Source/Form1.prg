/*
 * Projeto: brwexcel
 * Arquivo: Form1.prg
 * Descrição: Exemplo de como Exportar DBBrowse para Excel com XlsxWriter
 * Autor: Fausto Di Creddo Trautwein
 * Data: 12-27-2016
 */

#include "Xailer.ch"

CLASS TForm1 FROM TForm

   COMPONENT oDSConsulta
   COMPONENT oBrwConsulta
   COMPONENT oBtnExporta

   METHOD CreateForm()
   METHOD DSConsultaCreate( oSender )
   METHOD FormShow( oSender )
   METHOD BtnExportaClick( oSender )

ENDCLASS

#include "Form1.xfm"

//------------------------------------------------------------------------------

METHOD DSConsultaCreate( oSender ) CLASS TForm1
LOCAL aStruct

aStruct:= {;
{ "codigo", "N", 10, 0 },;
{ "nome", "C", 50, 0 },;
{ "valor", "N", 14, 2 } }
oSender:GetStructFrom(aStruct)
oSender:Open()

RETURN Nil

//------------------------------------------------------------------------------

METHOD FormShow( oSender ) CLASS TForm1
LOCAL i, nTot:= 0
::oDSConsulta:SaveState(.t.)
::oBrwConsulta:lRedraw:= .f.
FOR i:= 1 TO 40
   ::oDSConsulta:AddNew()
   ::oDSConsulta:codigo:= i
   ::oDSConsulta:nome:= "NOME DE TESTE DA PESSOA EXEMPLO "+ALLTRIM(STR(i))
   ::oDSConsulta:valor:= i * 1000
   nTot+= ::oDSConsulta:valor
   ::oDSConsulta:Update()
NEXT
::oDSConsulta:ReleaseState(.t.)
::oDSConsulta:GoTop()
::oBrwConsulta:lRedraw:= .t.

WITH OBJECT ::oBrwConsulta:ColWithHeader("Valor")
   :cFooter:= TRAN(nTot, :cPicture )
END

::oBrwConsulta:Refresh()
RETURN Nil

//------------------------------------------------------------------------------

METHOD BtnExportaClick( oSender ) CLASS TForm1
LOCAL aTitulo, aParametros:= {}
aTitulo:= { "Exemplo de Browse / Excel" }
AADD( aParametros, { "Teste com cores, picture e filtro." } )
GeraPlanilhaBrowsePython( ::oBrwConsulta, aTitulo, aParametros )
RETURN Nil

//------------------------------------------------------------------------------
