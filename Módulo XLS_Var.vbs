Option Compare Text
Option Explicit

'
'Define as variáveis que serão usadas'
'

'Define a operação a ser feita
Public ValidAtual As String

'Define a tabela a ser validada
Public Planilha As Worksheet

'Define a tabela a ser validada
Public Tabela As ListObject



'Define o threshold de dias que configurar atraso
Public Dias As Integer

'Define variável que acompanhará o Report das recomendações
Public Reporte As String

'Define variável para contar número de linhas
Public Linhas As Integer

'Define variável para contar número de colunas
Public Colunas As Integer

'Define i, variável que conta as linhas
Public i As Integer

'Contador de inconsistências dada numa dada linha
Public InconsistCount As Integer

'Contador de inconsistências dada numa dada linha
Public AlertType As String

'Define os parâmetros para o comentário gerencial da recomendação
Public ComentGer As String

'Define os parâmetros para a situação da recomendação
Public Situação As String

'Define os parâmetros para o status da recomendação
Public Status As String

'Define os parâmetros para a data inicial da recomendação
Public DataInicial As Date

'Define os parâmetros para a data planejada da recomendação
Public DataPlan As Date

'Define os parâmetros para a justificativa de replanejada da recomendação
Public JustifData As String

'Define os parâmetros para a a data replanejada da recomendação
Public DataReplan As Date

'Define os parâmetros para a data final da recomendação
Public DataFinal As Date