#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"

// --------------------------------------------------------------------------------
// Declaracao da Classe DataFrame
// --------------------------------------------------------------------------------

CLASS DataFrame  

// Declaracao das propriedades da Classe
DATA aCabecalho //lista com os cabecalhos do dataframe
DATA aDados     //matriz com os retornos do dataframe
DATA cQuery     //string com a query de consulta

// Declaração dos Métodos da Classe
METHOD New(cQuery) CONSTRUCTOR     //cria o objeto com os dados da query

ENDCLASS  
 
// Criação do construtor, onde atribuimos os valores default 
// para as propriedades e retornamos Self
METHOD New(cQuery) Class DataFrame
    ::aCabecalho  := {}
    ::aDados      := {} 
    ::cQuery      := cQuery 
    TCQUERY cQuery NEW ALIAS "DataFrame"
    DataFrame->(DBGotop()) 
    aCabecalho := DataFrame->(DbStruct())

    While DataFrame->(!EOF()) 

            aAux := {}

            For i := 1 to len(aCabecalho)
                cCampoAtu := aCabecalho[i][1] //percorre todos as colunas da query e coloca num vetor
                aadd(aAux,&("DataFrame->"+cCampoAtu)) 
            Next
            
            aadd(aDados, aClone(aAux))  
                   
        Dbskip()  

    Enddo 
    DbCloseArea()  

Return Self

 