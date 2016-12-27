/*
 * Projeto: brwexcel
 * Arquivo: Module1.prg
 * Descrição: Exportar DBBrowse para Excel
 * Autor: Fausto Di Creddo Trautwein
 * Data: 12-27-2016
 */

#include "Xailer.ch"
#DEFINE pCRLF  CHR(13)+CHR(10)

//-----------------
FUNCTION GeraPlanilhaBrowsePython( oBrwConsulta, aTitulo, aParametros, aWorkSheet )
//-----------------
aWorkSheet:= IIF( aWorkSheet == Nil, { "ws1", "Planilha1" }, aWorkSheet )
GeraPlanilhaBrowseSheet( { { oBrwConsulta, aTitulo, aParametros, aWorkSheet } } )
RETURN Nil

//-----------------
FUNCTION GeraPlanilhaBrowseSheet( aSheet )
//-----------------
LOCAL oXls, i

oXls:= XlsxWriter():New()

FOR i:= 1 TO LEN( aSheet )
   GeraPlanilhaBrowsePythonProcessa( oXls, aSheet[i,1], aSheet[i,2], aSheet[i,3], aSheet[i,4] )
NEXT

oXls:SaveWorkBook()
oXls:RunPython()
oXls:ShowWorkBook()
RETURN Nil

//-----------------
FUNCTION GeraPlanilhaBrowsePythonProcessa( oXls, oBrwConsulta, aTitulo, aParametros, aWorkSheet )
//-----------------
LOCAL nLinha, oRange, nColunas, i, nKeyNoAnt, aColunas, cLinha, nColuna, xvalor
LOCAL cPicture, cPict1, cPict2, cNumberFormat, oCol, xvalorret, nClrPane, nLinhaInicial
LOCAL hPicture:= { => }, lAlt:= .t.
LOCAL nColunaInicial, nLinhaTop, nClrText, nKeyNo, lHighLite, oCell, oInterior, oCell1, oCell2, nImage
LOCAL cPython, cAux, nCol, aCol:= { => }, cFormato, aFormato
LOCAL nBookMark

oXls:Add_Worksheet( aWorkSheet[1], aWorkSheet[2] )

nKeyNo:= oBrwConsulta:KeyNo()
nBookMark:= oBrwConsulta:BookMark()
oBrwConsulta:lRedraw:= .f.

aTitulo    := IIF( aTitulo     == Nil, {}, aTitulo     )
aParametros:= IIF( aParametros == Nil, {}, aParametros )

nColunaInicial:= -1
nLinhaTop:= 0 //nLinhaPlanilha + 1
nColunas:= LEN( oBrwConsulta:aCols )

// obtem as colunas visíveis e a ordem em que é mostrada no browse
nCol:= 0
FOR i:= 1 TO nColunas
   IF oBrwConsulta:aCols[i]:lVisible
      aCol[++nCol]:= i
   ENDIF
NEXT

// Mostra o título
nLinha:= nLinhaTop
aFormato:= { => }
aFormato[ 'font_size' ]:= '16'
aFormato[ 'bold'      ]:= 'True'
aFormato[ 'bg_color'  ]:= '"#F5F5F5"'
aFormato[ 'align'     ]:= '"center"'
aFormato[ 'border'    ]:= '1'
aFormato[ 'valign'     ]:= '"center"'
cFormato:= oXls:SetFormat( aFormato )

FOR i:= 1 TO LEN( aTitulo )
   oXls:Set_Row( nLinha, 20 )
   oXls:Merge( nLinha, 1 + nColunaInicial, nLinha, LEN( aCol ) + nColunaInicial, aTitulo[i], cFormato)
   nLinha++
NEXT

// Mostra os parâmetros
aFormato:= { => }
aFormato[ 'bold'      ]:= 'True'
aFormato[ 'bg_color'  ]:= '"#F5F5F5"'
aFormato[ 'border'    ]:= '1'
cFormato:= oXls:SetFormat( aFormato )
FOR i:= 1 TO LEN( aParametros )
   oXls:Merge( nLinha, 1 + nColunaInicial, nLinha, LEN( aCol ) + nColunaInicial, IIF( VALTYPE( aParametros[i] ) == "A", aParametros[i,1], aParametros[i]), cFormato)
   nLinha++
NEXT

nLinhaInicial:= nLinha:= nLinhaTop + LEN( aTitulo ) + LEN( aParametros )

oXls:Set_Row( nLinha, oBrwConsulta:nHeaderHeight )

FOR i:= 1 TO LEN( aCol )
   aFormato:= { => }
   IF oBrwConsulta:aCols[aCol[i]]:nAlignment == taRIGHT;
      .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taLEFTHEADERRIGHT;
      .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taCENTERHEADERRIGHT
      aFormato[ "align" ]:= '"right"'
   ELSEIF oBrwConsulta:aCols[aCol[i]]:nAlignment == taCENTER;
      .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taLEFTHEADERCENTER;
      .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taRIGHTHEADERCENTER
      aFormato[ "align" ]:= '"center"'
   ENDIF

   aFormato[ "bold"      ]:= 'True'
   aFormato[ "bg_color"  ]:= '"#A5A5A5"'
   aFormato[ "text_wrap" ]:= 'True'
   aFormato[ "border"    ]:= '1'

   cFormato:= oXls:SetFormat(  aFormato )

   oXls:Write( nLinha, i + nColunaInicial, STRTRAN(oBrwConsulta:aCols[aCol[i]]:cHeader,pCRLF," "), cFormato )
   oXls:Set_Column( i + nColunaInicial, i + nColunaInicial, oBrwConsulta:aCols[aCol[i]]:nWidth*0.16 )
NEXT

nLinha++

oBrwConsulta:GoTop()

FOR i:= 1 TO LEN( aCol )
   xvalor:= oBrwConsulta:aCols[aCol[i]]:GetData() //Value()
   cPicture:= oBrwConsulta:aCols[aCol[i]]:cPicture
   hPicture[ i ]:= ""
   IF VALTYPE( xvalor ) == "N"
      IF !EMPTY(cPicture)
         IF LEFT( cPicture,1 ) == "@" // retira o @E
            cPicture:= SUBS( cPicture, HB_AT(" ",cPicture )+1, LEN(cPicture) )
         ENDIF
         cPicture:= STRTRAN( cPicture, "9", "#" )
         // 999.999,99
         cPict1:= LEFT( cPicture, RAT(".",cPicture) - 2 )
         cPict2:= SUBS( cPicture, RAT(".",cPicture) - 1, LEN(cPicture ) )
         cPict2:= STRTRAN( cPict2, "#", "0" )
         cNumberFormat:= cPict1+cPict2
         hPicture[ i ]:= cNumberFormat
      ENDIF
   ENDIF
NEXT

WHILE .t.

   FOR i:= 1 TO LEN( aCol )
      oCol:= oBrwConsulta:aCols[aCol[i]]
      xvalor:= oCol:Value
      nImage:= 0
      IF  oCol:EventAssigned( "OnGetData" )
         xvalorret:= oCol:OnGetData(@xvalor,@nImage)
         IF !(xvalorret == Nil)
            IF EMPTY( xvalorret ) .and. !EMPTY(nImage)
            ELSE
               xvalor:= xvalorret
            ENDIF
         ENDIF
      ENDIF

      nClrPane:= oCol:nClrPane
      nClrText:= oCol:nClrText
      lHighLite:= .f.
      IF  oCol:EventAssigned( "OnDrawCell" )
         oCol:OnDrawCell(ValueToString(xvalor), @nClrText, @nClrPane, lHighLite)
      ENDIF
         //nClrPane:= oCol:nClrPane
         IF nClrPane == 16777215 // 0xFFFFFF
            IF lAlt
               nClrPane:= oBrwConsulta:nClrAltPane
            ELSE
               nClrPane:= oBrwConsulta:nClrPane
            ENDIF
         ENDIF
      //ENDIF

      aFormato:= { => }
      IF oBrwConsulta:aCols[aCol[i]]:nAlignment == taRIGHT;
         .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taRIGHTHEADERCENTER;
         .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taRIGHTHEADERLEFT
         aFormato["align"]:= '"right"'
      ELSEIF oBrwConsulta:aCols[aCol[i]]:nAlignment == taCENTER;
         .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taCENTERHEADERRIGHT;
         .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taCENTERHEADERLEFT
         aFormato["align"]:= '"center"'
      ENDIF

      aFormato[ "font_color" ]:= '"#'+NumToCor(nClrText)+'"'
      IF !EMPTY(hPicture[ i ])
         aFormato[ "num_format" ]:= '"'+hPicture[i]+'"'
      ENDIF
      aFormato[ "bg_color"   ]:= '"#'+NumToCor(nClrPane)+'"'
      aFormato[ "border"     ]:=  '1'

      cFormato:= oXls:SetFormat( aFormato )

      oXls:Write( nLinha, i + nColunaInicial, xvalor, cFormato )

   NEXT
   lAlt:= !lAlt

   nLinha++

   // posição antes de tentar ir para a próxima linha
   nKeyNoAnt:= oBrwConsulta:KeyNo()

   oBrwConsulta:GoDown()

   // verifica se mudou para a próxima linha, se não então chegou ao fim
   IF oBrwConsulta:KeyNo() == nKeyNoAnt
      EXIT
   ENDIF

ENDDO

IF oBrwConsulta:lFooter
   FOR i:= 1 TO LEN( aCol )
      aFormato:= { => }
      IF oBrwConsulta:aCols[aCol[i]]:nAlignment == taRIGHT;
         .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taRIGHTHEADERCENTER;
         .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taRIGHTHEADERLEFT
         aFormato["align"]:= '"right"'
      ELSEIF oBrwConsulta:aCols[aCol[i]]:nAlignment == taCENTER;
         .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taCENTERHEADERRIGHT;
         .or. oBrwConsulta:aCols[aCol[i]]:nAlignment == taCENTERHEADERLEFT
         aFormato["align"]:= '"center"'
      ENDIF

      aFormato[ "bold"  ]:= 'True'
      aFormato[ "border"]:=  '1'
      aFormato[ "text_wrap" ]:= 'True'

      cFormato:= oXls:SetFormat( aFormato )
      oXls:Write( nLinha, i + nColunaInicial, oBrwConsulta:aCols[aCol[i]]:cFooter, cFormato )
   NEXT
   nLinha++
ENDIF

oBrwConsulta:lRedraw:= .t.
oBrwConsulta:BookMark(nBookMark)

RETURN Nil

//-----------------
Function ValueToString(xField)
//-----------------
Local cType, result := nil

cType := ValType(xField)

if cType == "D"
        IF !EMPTY(xField)
           result := dtos( xField )
        ENDIF
elseif cType == "N"
        result :=ALLTRIM(str(xField))
elseif cType == "L"
        result := iif( xField, "t", "f" )
elseif cType == "C" .or. cType == "M"
        result := xField
end
return result

//-----------------
FUNCTION NumToCor(nCor)
//-----------------
LOCAL cRet
cRet:= HB_NumToHex(nCor,6)
cRet:= RIGHT( cRet, 2 ) + SUBS( cRet, 3, 2 ) + LEFT( cRet, 2 )
RETURN cRet

//-------------------
FUNCTION CurrentDir()
//-------------------
LOCAL cDir
cDir:= Application:CurrentDir
/*
Type: C >>>A:\\192.168.1.248\Latic\FVendas2\mapa.html<<<
*/
IF SUBS( cDir, 2, 3 ) == ":\\" //LEFT( cDir, 2 ) == "\:"
   cDir:= SUBS( cDir, 3, LEN( cDir ) )
ENDIF
RETURN cDir
