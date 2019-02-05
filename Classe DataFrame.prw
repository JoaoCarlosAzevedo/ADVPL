#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"

// --------------------------------------------------------------------------------
// Declaracao da Classe DataFrames
// --------------------------------------------------------------------------------

CLASS DataFrames FROM TReport  

// Declaracao das propriedades da Classe
DATA aCabecalho //lista com os cabecalhos do dataframe
DATA aDados     //matriz com os retornos do dataframe
DATA cQuery     //string com a query de consulta
DATA oRelatorio //Objeto do tipo TReport 

// Declaração dos Métodos da Classe
METHOD New(cQuery) CONSTRUCTOR     //cria o objeto com os dados da query
METHOD Relatorio()                 //Metodo retorna um objeto do tipo TReport ja configurado.
METHOD Calc(cOperador,nPivo,nAlvo)  //Metodo que retorna array com duas dimensoes conforme parametro solicitado
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
                aadd(aAux,&("DataFrame->"+cCampoAtu)) 
            Next
            
            aadd(::aDados , aClone(aAux))  
                   
        Dbskip()  
 
    Enddo 
    DbCloseArea()  

Return Self


METHOD CALC(cOperador,nPivo,nAlvo) Class DataFrames
	Local aRet := {} 
	Local nPos := 0
	
	For i := 1 to Len(::aDados)
	
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

	Next

Return aRet
 
METHOD Relatorio() Class DataFrames 

    	//Criação do componente de impressão

	::oRelatorio := TReport():New(	"MovD3",;		//Nome do Relatório
								"Relatorio de Movimentacoes",;		//Título
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
        //TRCell():New(oSectDad, "TIPOMOV", "QRY_AUX", "Tipomov", /*Picture*/, 3, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
        //TRCell():New(oSectDad, "CODIGO", "QRY_AUX", "Codigo", /*Picture*/, 15, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
        //TRCell():New(oSectDad, "DESCRICAO", "QRY_AUX", "Descricao", /*Picture*/, 30, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
        //TRCell():New(oSectDad, "UNIDADE", "QRY_AUX", "Unidade", /*Picture*/, 2, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
        //TRCell():New(oSectDad, "QUANTIDADE", "QRY_AUX", "Quantidade", /*Picture*/, 15, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
    Next

    //::oRelatorio:PrintDialog()
 
Return ::oRelatorio  
 
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
		//Imprimindo a linha atual
		//oSectDad:PrintLine()
        For y := 1 to Len(aCabec)
            oSectDad:Cell(aCabec[y,1]):SetValue(aDados[i,y])
            oSectDad:Cell(aCabec[y,1]):SetAlign("LEFT")
        Next 
        oSectDad:PrintLine()
	Next
	oSectDad:Finish()
  
Return 
  