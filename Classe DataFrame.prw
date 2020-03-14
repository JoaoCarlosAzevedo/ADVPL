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
METHOD Excel(cTile1,cTitle2,cFileName,cDir)
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