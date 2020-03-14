# ADVPL - Classe DataFrame

Classe desenvolvida em ADVPL para facilitar a manipulação de Querys, construção de relatórios, analise de dados, construção de gráficos. 

## Atributos da Classe
* aCabecalho - Array com as colunas resultantes da query. [ X ]
* aDados     - Array (matriz) com os dados resultantes da query. [ X ] 
* cQuery     - String com a query que foi passada no método do construtor. [ X ]
* oRelatorio - Objeto do tipo TReport  [ X ]

## Métodos
* New(cQuery) - Construtor onde o parametro é a query que deseja construir o objeto. [ X ]
* Relatorio() - Instancia classe do TReport no atributo oRelatorio para geração de relatorios [ X ]
* Excel(cTituloPlanilha,cTituloTabela,cNomeArquivo,cDiretorio) - Cria excel conforme os parametros informados [ X ]
* Soma  - [ X ] 
* Media - [ X ] 
* Plot  - [   ] Pendente
* Excel - [ X ] 
