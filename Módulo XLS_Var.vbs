Option Compare Text
Option Explicit

'
'Define as vari�veis que ser�o usadas'
'

'Define a opera��o a ser feita
Public ValidAtual As String

'Define a tabela a ser validada
Public Planilha As Worksheet

'Define a tabela a ser validada
Public Tabela As ListObject



'Define o threshold de dias que configurar atraso
Public Dias As Integer

'Define vari�vel que acompanhar� o Report das recomenda��es
Public Reporte As String

'Define vari�vel para contar n�mero de linhas
Public Linhas As Integer

'Define vari�vel para contar n�mero de colunas
Public Colunas As Integer

'Define i, vari�vel que conta as linhas
Public i As Integer

'Contador de inconsist�ncias dada numa dada linha
Public InconsistCount As Integer

'Contador de inconsist�ncias dada numa dada linha
Public AlertType As String

'Define os par�metros para o coment�rio gerencial da recomenda��o
Public ComentGer As String

'Define os par�metros para a situa��o da recomenda��o
Public Situa��o As String

'Define os par�metros para o status da recomenda��o
Public Status As String

'Define os par�metros para a data inicial da recomenda��o
Public DataInicial As Date

'Define os par�metros para a data planejada da recomenda��o
Public DataPlan As Date

'Define os par�metros para a justificativa de replanejada da recomenda��o
Public JustifData As String

'Define os par�metros para a a data replanejada da recomenda��o
Public DataReplan As Date

'Define os par�metros para a data final da recomenda��o
Public DataFinal As Date