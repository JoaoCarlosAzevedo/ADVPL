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

// Declara��o dos M�todos da Classe
METHOD New(cQuery) CONSTRUCTOR     //cria o objeto com os dados da query
METHOD Relatorio()                 //Metodo retorna um objeto do tipo TReport ja configurado.
ENDCLASS    

 
// Cria��o do construtor, onde atribuimos os valores default 
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
 
METHOD Relatorio() Class DataFrames 

    	//Cria��o do componente de impress�o

	::oRelatorio := TReport():New(	"MovD3",;		//Nome do Relat�rio
								"Relatorio de Movimentacoes",;		//T�tulo
								,;		//Pergunte ... Se eu defino a pergunta aqui, ser� impresso uma p�gina com os par�metros, conforme privil�gio 101
								{|oReport| fRepPrint(::oRelatorio,::aDados,::aCabecalho)},;		//Bloco de c�digo que ser� executado na confirma��o da impress�o
								)		//Descri��o
	::oRelatorio:SetTotalInLine(.F.)
	::oRelatorio:lParamPage := .F.
	::oRelatorio:oPage:SetPaperSize(9) //Folha A4 
	::oRelatorio:SetPortrait()
	
	//Criando a se��o de dados
	oSectDad := TRSection():New(	::oRelatorio,;		//Objeto TReport que a se��o pertence
									"Dados",;		//Descri��o da se��o
									{"QRY_AUX"})		//Tabelas utilizadas, a primeira ser� considerada como principal da se��o
	oSectDad:SetTotalInLine(.F.)  //Define se os totalizadores ser�o impressos em linha ou coluna. .F.=Coluna; .T.=Linha
	
	//Colunas do relat�rio
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
	
	//Pegando as se��es do relat�rio
	oSectDad := oReport:Section(1)
	
	nTotal := LEN(aDados)

	oReport:SetMeter(nTotal)
	
	//Enquanto houver dados
	oSectDad:Init()

	//Incrementando a r�gua
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
  