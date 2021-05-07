#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#include "TOTVS.CH"

// --------------------------------------------------------------------------------
// Declaracao da Classe DataFrames
// --------------------------------------------------------------------------------
CLASS DataFrames FROM TReport  

// Declaracao das propriedades da Classe
DATA aCabecalho //lista com os cabecalhos do dataframe
DATA aDados     //matriz com os retornos do dataframe
DATA cQuery     //string com a query de consulta 
DATA oRelatorio //Objeto do tipo TReport 
DATA oBrowse
DATA oFwBrowse
 
// Declaração dos Métodos da Classe
METHOD New(cQuery) CONSTRUCTOR               //cria o objeto com os dados da query
METHOD Relatorio()                           //Metodo retorna um objeto do tipo TReport ja configurado.
METHOD Calc(cOperador,nPivo,nAlvo)           //Metodo que retorna array com duas dimensoes conforme parametro solicitado
METHOD Excel(cTile1,cTitle2,cFileName,cDir)
METHOD Browse(oDialog,bAction)
METHOD isEmpty()
METHOD getValue(nLine,cField)
METHOD FwBrowse(oPanel)

ENDCLASS    

 
// Criação do construtor, onde atribuimos os valores default 
// para as propriedades e retornamos Self
METHOD New(cQuery) Class DataFrames
    ::aCabecalho  := {}
    ::aDados      := {}   
    ::cQuery      := cQuery 
    TCQUERY cQuery NEW ALIAS "DataFrame"
    DataFrame->(DBGotop()) 
    ::aCabecalho := DataFrame->(DbStruct())

    While DataFrame->(!EOF()) 

            aAux := {}

            For i := 1 to len(::aCabecalho)
                cCampoAtu := ::aCabecalho[i][1] //percorre todos as colunas da query e coloca num vetor
				IF ALLTRIM(::aCabecalho[i][1]) == 'CHECKBOX'
					aadd(aAux,.F.) 
				ELSE
					aadd(aAux,&("DataFrame->"+cCampoAtu)) 
				ENDIF 
                	
            Next  
            
            aadd(::aDados , aClone(aAux))  
                   
        Dbskip()  
 
    Enddo 
    DbCloseArea()  

Return Self


METHOD CALC(cOperador,nPivo,nAlvo) Class DataFrames
	Local aRet := {} 
	Local nPos := 0
	Local nSum := 0

	For i := 1 to Len(::aDados)

		IF nPivo == 0 //nao tem regra no agrupamento
			nSum := nSum + ::aDados[i,nAlvo]
		ELSE	
			nPos := aScanX( aRet,{ |X,Y| X[1] == ::aDados[i,nPivo]} )		
			IF nPos == 0 //nao encontrou
				DO CASE 
					CASE cOperador == "SOMA"
						Aadd(aRet,{::aDados[i,nPivo],::aDados[i,nAlvo]})
					CASE cOperador == "MEDIA"
						Aadd(aRet,{::aDados[i,nPivo],::aDados[i,nAlvo],1}) //adiciona uma dimensao a mais que ira fazer a contagem
					CASE cOperador == "CONTAGEM"
						Aadd(aRet,{::aDados[i,nPivo],1}) 
				ENDCASE
			ELSE 
				DO CASE
					CASE cOperador == "SOMA"
						aRet[nPos,2] := aRet[nPos,2] + ::aDados[i,nAlvo] //realiza a soma conforme posicao
					CASE cOperador == "MEDIA"
						aRet[nPos,2] := aRet[nPos,2] + ::aDados[i,nAlvo]
						aRet[nPos,3] := aRet[nPos,3] + 1 //soma com mais um para depois fazer divis?o
					CASE cOperador == "CONTAGEM"
						aRet[nPos,2] := aRet[nPos,2] + 1
				ENDCASE
			ENDIF
		ENDIF
	Next

	IF nPivo == 0
		DO CASE 
			CASE cOperador == "SOMA"
				return nSum
			CASE cOperador == "MEDIA"
				return nSum / LEN(::aDados)
			CASE cOperador == "CONTAGEM"
				return LEN(::aDados)
		ENDCASE	
	ENDIF
 
Return aRet
 
METHOD Relatorio() Class DataFrames 

    	//Criação do componente de impressão

	::oRelatorio := TReport():New(	"DataFrame",;		//Nome do Relatório
								"Relatorio de Dados",;		//Título
								,;		//Pergunte ... Se eu defino a pergunta aqui, será impresso uma página com os parâmetros, conforme privilégio 101
								{|oReport| fRepPrint(::oRelatorio,::aDados,::aCabecalho)},;		//Bloco de código que será executado na confirmação da impressão
								)		//Descrição
	::oRelatorio:SetTotalInLine(.F.)
	::oRelatorio:lParamPage := .F.
	::oRelatorio:oPage:SetPaperSize(9) //Folha A4 
	::oRelatorio:SetPortrait()
	 
	//Criando a seção de dados
	oSectDad := TRSection():New(	::oRelatorio,;		//Objeto TReport que a seção pertence
									"Dados",;		//Descrição da seção
									{"QRY_AUX"})		//Tabelas utilizadas, a primeira será considerada como principal da seção
	oSectDad:SetTotalInLine(.F.)  //Define se os totalizadores serão impressos em linha ou coluna. .F.=Coluna; .T.=Linha
	
	//Colunas do relatório
    For i := 1 to Len(::aCabecalho) 
        
        TRCell():New(oSectDad, ::aCabecalho[i][1], "QRY_AUX",::aCabecalho[i][1], /*Picture*/, ::aCabecalho[i][3], /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)

   Next
 
Return ::oRelatorio  

METHOD isEmpty() Class DataFrames 
	lRet := .F.

	IF EMPTY(LEN(::aDados))
		lRet := .T.
	ENDIF
 
Return lRet

METHOD getValue(nLine,cField) Class DataFrames 
	Local nIndex := 0
	
	nIndex :=  aScanX(::aCabecalho,{ |X,Y| X[1] == ALLTRIM(UPPER(cField)) })	

	IF nIndex <> 0 
		return ::aDados[nLine,nIndex]
	ENDIF  
   
Return ""
 
Static Function fRepPrint(oReport,aDados,aCabec) 
	Local aArea    := GetArea()
	Local cQryAux  := "" 
	Local oSectDad := Nil
	Local nAtual   := 0
	Local nTotal   := 0
	
	//Pegando as seções do relatório
	oSectDad := oReport:Section(1)
	
	nTotal := LEN(aDados)

	oReport:SetMeter(nTotal)
	
	//Enquanto houver dados
	oSectDad:Init()

	//Incrementando a régua
    For i := 1 to len(aDados)
		nAtual++
		oReport:SetMsgPrint("Imprimindo registro "+cValToChar(nAtual)+" de "+cValToChar(nTotal)+"...")
		oReport:IncMeter()
        For y := 1 to Len(aCabec)
            oSectDad:Cell(aCabec[y,1]):SetValue(aDados[i,y])
            oSectDad:Cell(aCabec[y,1]):SetAlign("LEFT")
        Next 
        oSectDad:PrintLine()
	Next
	oSectDad:Finish()
  
Return 

METHOD Excel(cTile1,cTitle2,cFileName,cDir) Class DataFrames
		Local i     := 0
		Local j     := 0
		Local aLine := {}
        Local oFWMsExcel 
        Local oExcel 
        Local cArquivo := cDir+cFileName 
        
        //verificar se existe o arquivo 
        IF FILE(cArquivo) 
            If FERASE(cArquivo) == -1
                MsgStop('Falha na deleção do Arquivo de Excel: ' + cArquivo + Chr(13)+Chr(10) + "Conferir se não esta aberto...")
                return .F.
            Endif 
        ENDIF
 
        //Criando o objeto que irá gerar o conteúdo do Excel
        oFWMsExcel := FWMSExcel():New()

        //Aba 01 - Teste
        oFWMsExcel:AddworkSheet(cTile1) //Não utilizar número junto com sinal de menos. Ex.: 1-

        //Criando a Tabela 
        oFWMsExcel:AddTable(cTile1,cTitle2) 
        
		//Criando Colunas
		For i := 1 to Len(::aCabecalho) 
			IF ::aCabecalho[i,2] == "N" //se o campo for numero coloca o como valor
				oFWMsExcel:AddColumn(cTile1,cTitle2  ,::aCabecalho[i,1],    1,2) //1 = Modo Texto //2 = Valor sem R$ //3 = Valor com R$
			ELSE	
				oFWMsExcel:AddColumn(cTile1,cTitle2  ,::aCabecalho[i,1],    1,1) //1 = Modo Texto //2 = Valor sem R$ //3 = Valor com R$
			ENDIF
		Next

		//Criando as linhas
		For i := 1 to Len(::aDados) 
			aLine := {}
			For j := 1 to Len(::aCabecalho)
				AADD( aLine, ::aDados[i,j])	
			Next 
			oFWMsExcel:AddRow(cTile1,cTitle2,aLine)
		Next 

    //Ativando o arquivo e gerando o xml 
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo) 
          
    //Abrindo o excel e abrindo o arquivo xml
    oExcel := MsExcel():New()               //Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArquivo)         //Abre uma planilha
    oExcel:SetVisible(.T.)                  //Visualiza a planilha
    oExcel:Destroy()                        //Encerra o processo do gerenciador de tarefas
     
return Self 

 
 METHOD Browse(oDialog,bAction) Class Dataframes
	Local x        := 0
	Local nSize    := 0
	Local cAlign   := ""
	Local lObject  := .F.
	Local cPicture := "" 
	Local bBuild   := {|| } 
	Local aAux     := {}
	Static oDlg 

	IF oDialog == nil 
		DEFINE MSDIALOG oDlg TITLE "PRINT" FROM 000, 000  TO 500, 500 COLORS 0, 16777215 PIXEL
		::oBrowse := TCBrowse():New(000,000,260,184,,,,oDlg,,,,,,,,,,,,.F.,"",.T.,,.F.,,,)
	ELSE
		::oBrowse := TCBrowse():New(000,000,260,184,,,,oDialog,,,,,,,,,,,,.F.,"",.T.,,.F.,,,)   	
	ENDIF
	
	
	::oBrowse:setArray( ::aDados )
	
  	For x := 1 to Len(::aCabecalho)
		nSize    := ::aCabecalho[x,3]
		cAlign   := AlignField(::aCabecalho[x,2]) 
		lObject  := .F. //::aCabecalho[i,2] - tipo for oBject? 
		cPicture := "" 
		bBuild   := &( "{ ||   self:aDados[self:oBrowse:nAt , " +CVALTOCHAR( x )+ "]  } " )  
 
		::oBrowse:AddColumn(TCColumn():New(::aCabecalho[x,1],bBuild,cPicture,,,cAlign,nSize,lObject,.T.,,,,,))  
  
	Next       
  
	::oBrowse:bLDblClick    := bAction     
    ::oBrowse:bHeaderClick := {|| Alert( cValtoChar(::oBrowse:nColPos) )}
	::oBrowse:Align := CONTROL_ALIGN_ALLCLIENT     

	IF oDialog == nil 
	 	ACTIVATE MSDIALOG oDlg CENTERED   
	ENDIF
			  
 return ::oBrowse 
  

METHOD FwBrowse(oDialog) Class Dataframes
	Local nSize      := 0 
	Local lObject    := .F.
	Local cPicture   := "" 
	Local bBuild     := {|| } 
	Local aSeek      := {}
	Local aFieFilter := {}
 
    ::oFwBrowse := fwBrowse():New() 
 
    ::oFwBrowse:setOwner( oDialog )  
 
    ::oFwBrowse:setDataArray()  
    ::oFwBrowse:setArray( ::aDados )  
 
    ::oFwBrowse:SetLocate() // Habilita a Localização de registros

    For i := 1 to Len( ::aCabecalho )   
		
		IF ALLTRIM(::aCabecalho[i,1]) == 'CHECKBOX'

			::oFwBrowse:AddMarkColumns({|| IIf(::aDados[::oFwBrowse:nAt,01], "LBOK", "LBNO")},; //Code-Block image
				{|| SelectOne(::oFwBrowse,    ::aDados)},; //Code-Block Double Click  
				{|| SelectAll(::oFwBrowse,    ::aDados) }) //Code-Block Header Click */

		ELSE

			IF isLegend(::aCabecalho[i,1]) 
				
				bBuild   := &( "{ || self:aDados[self:oFwBrowse:nAt," +CVALTOCHAR( i )+ " ] } "  )  

				::oFwBrowse:addColumn( {'',;
										bBuild ,; 
										::aCabecalho[i,2],;  
										'@!',;
										0,;
										1,;
										,;
										.T. ,;
										,;  
										.T.,;
										,; 
										"self:aDados[self:oFwBrowse:nAt," +CVALTOCHAR( i )+ "]",;
										,; 
										.F.,; 
										.T.,; 
										,::aCabecalho[i,1]}) 
			ELSE
				nSize    := ::aCabecalho[i,3]  
				nAlign   := AlignFFw(::aCabecalho[i,2])   
				lObject  := .F.   
				cPicture := cRetPicture( ::aCabecalho[i,2], ::aCabecalho[i,4]  ) 
				bBuild   := &( "{ || self:aDados[self:oFwBrowse:nAt," +CVALTOCHAR( i )+ " ] } "  )  

				::oFwBrowse:addColumn( {::aCabecalho[i,1] ,;
									bBuild ,;
									::aCabecalho[i,2],; 
									cPicture,;
									nAlign,;
									nSize,;
									,;
									.T. ,;
									,;  
									.F.,;
									,;
									"self:aDados[self:oFwBrowse:nAt," +CVALTOCHAR( i )+ "]",;
									,; 
									.F.,; 
									.T.,; 
									,::aCabecalho[i,1]  }) 
			 
				IF ::aCabecalho[i,2] == 'C'                       
					Aadd(aSeek,{::aCabecalho[i,1],      {{"","C",nSize,0, "self:aCabecalho[i,1]" ,"@!"     }}, i, .T. } )
				ENDIF 
			ENDIF 
			Aadd(aFieFilter,{::aCabecalho[i,1],::aCabecalho[i,1],::aCabecalho[i,2], nSize, 0, cPicture}) 
		ENDIF
    Next     
 
    //::oFwBrowse:setEditCell( .T. , { ||  ,.T.  } ) //activa edit and code block for validation
 
    ::oFwBrowse:SetSeek(nil,aSeek)   
    ::oFwBrowse:SetUseFilter()  
    ::oFwBrowse:SetFilterDefault( "" ) 
    ::oFwBrowse:SetFieldFilter(aFieFilter)  
 
 
    //::oFwBrowse:Activate(.T.) 
    
 return ::oFwBrowse  

 
Static Function SelectOne(oBrowse, aArquivo)
    aArquivo[oBrowse:nAt,1] := !aArquivo[oBrowse:nAt,1]
    //oBrowse:Refresh()
Return .T. 
  

Static Function SelectAll(oBrowse, aArquivo) 
	Local _ni := 1

	For _ni := 1 to len(aArquivo)
		aArquivo[_ni,1] := !aArquivo[_ni,1]
	Next
	oBrowse:Refresh() 

	//lMarker:=!lMarker 

Return .T.



Static function isLegend(cCampo)
	Local lRet := .F.

	IF Substring(cCampo,1,7) == 'LEGENDA'
		lRet := .T.
	ENDIF

return lRet  

Static Function cRetPicture(cTipo, nDecimal)
    Local cRet := ""
    Local nCasa := IIF(nDecimal==8,2,nDecimal)
  
    IF cTipo == "N"
        cRet :=  "@E 999,999,999." + Replicate( '9', nCasa )
    ELSE
        cRet := "@!"
    ENDIF   
 
return cRet

  Static Function AlignField(cTipo)
	Local nRet := 0
 
	DO CASE
		CASE cTipo == "N"
			nRet := 0
		CASE cTipo == "C"
			nRet := 1
		CASE cTipo == "D"
		    nRet := 2
		CASE cTipo == "O" 
		    nRet := 0 
		OTHERWISE
			nRet := 0 
	ENDCASE

 Return nRet

 Static Function AlignFFw(cTipo)
	Local cRet := ""   
 
	DO CASE
		CASE cTipo == "N"
			cRet := "RIGHT" 
		CASE cTipo == "C"
			cRet := "LEFT"  
		CASE cTipo == "D"
			cRet := "CENTER"
		CASE cTipo == "O" 
			cRet := "CENTER"	
		OTHERWISE
			cRet := "LEFT"
	ENDCASE

 Return cRet
 
 
User Function TestData()                        
Local oPanel1
Local oPanel2
Local cQuery := ""
Local oDados
local oDados2
Local bAction := {|| Alert("Teste") }
Local oTcBrowse
Local oTcBrowse2

Static oDlg

	cQuery +=" Select Top 5 C2_NUM, C2_ITEM, C2_SEQUEN, C2_DPROD, C2_QUANT, C2_QUANT * 10.333 as Quant2, CONVERT(CHAR,Cast(C2_DATPRI as date),103) as Data   "
	cQuery +=" From SC2010 "
	cQuery +=" Where SC2010.D_E_L_E_T_ <> '*'  "

	oDados := DataFrames():New(cQuery)  
	//oDados2 := DataFrames():New(cQuery)   

  DEFINE MSDIALOG oDlg TITLE "New Dialog" FROM 000, 000  TO 500, 500 COLORS 0, 16777215 PIXEL

    @ 000, 000 MSPANEL oPanel1 SIZE 250, 064 OF oDlg COLORS 0, 16777215 RAISED
    @ 064, 000 MSPANEL oPanel2 SIZE 250, 185 OF oDlg COLORS 0, 16777215 RAISED
	
	oTcBrowse  := oDados:FwBrowse(oPanel1) 

	//oTcBrowse2 := oDados2:Browse(nil,bAction)
  
   
	//oDados:Browse(oPanel2,bAction)
     
    // Don't change the Align Order    
    oPanel1:Align := CONTROL_ALIGN_TOP 
    oPanel2:Align := CONTROL_ALIGN_ALLCLIENT     
  
  ACTIVATE MSDIALOG oDlg CENTERED  

Return   

User Function Test02()
	Local cQuery := ""
	Local oDados2
	Local oTcBrowse2
	Local bAction := {|| Alert("Teste") } 

	cQuery +=" Select Top 5 C2_NUM, C2_ITEM, C2_SEQUEN, C2_DPROD, C2_QUANT, C2_QUANT * 10.333 as Quant2, CONVERT(CHAR,Cast(C2_DATPRI as date),103) as Data   "
	cQuery +=" From SC2010 "
	cQuery +=" Where SC2010.D_E_L_E_T_ <> '*'  "

	oDados2 := DataFrames():New(cQuery)   
	oTcBrowse2 := oDados2:Browse(nil,bAction)
Return  
