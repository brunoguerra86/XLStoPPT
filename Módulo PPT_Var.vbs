'Definir as vari�veis que ser�o usadas no PPT:
Public appPPT As PowerPoint.Application 'Abre um programa Powerpoint

Public prsntPPT As PowerPoint.Presentation 'Cria Apresenta��o de Powerpoint

Public slidePPT As PowerPoint.Slide 'Slides

Public shapePPT As PowerPoint.Shape 'Adiciona uma forma

Public headerPPT As PowerPoint.Shape 'Adiciona o cabe�alho

Public headerlinePPT As PowerPoint.Shape 'Adiciona o cabe�alho

Public SlideCountPPT As Long 'Conta slide

Public PathPPT As String, NamePPT As String 'Path at� o PPT

Public Total_RowCountPPT As Long 'quantas linhas de recomenda��o tem

Public RowCountPPT As Long 'quantas linhas o slide tem

'Define o primeiro n�vel de filtro do report
Public Filter1 As String

'Define as entradas poss�veis desse filtro
Public Op��esFilter1 As Collection
Public vnum1 As Variant

'Define por o que o filtro 1 foi filtrado
Public FilterAs1 As String

'Define o segundo n�vel de filtro do report
Public Filter2 As String

'Define as entradas poss�veis desse segundo filtro
Public Op��esFilter2 As Collection
Public vnum2 As Variant

'Define por o que o filtro 2 foi filtrado
Public FilterAs2 As String

'Define quais linhas sobram depois do filtro
Public VisibleRow As Range

'Define o primeiro n�vel de ordena��o do report
Public Class1 As String

'Define o segundo n�vel de ordena��o do report
Public Class2 As String

'Define o terceiros n�vel de ordena��o do report
Public Class3 As String

'Essas determinam onde vai buscar cada coluna do PPT
Public Coluna1 As String
Public Coluna2 As String
Public Coluna3 As String
Public Coluna4 As String
Public Coluna5 As String
Public Coluna6 As String
Public Coluna7 As String
Public Coluna8 As String

'Acompanham os loops do class1 durante o preenchimento da tabela
Public ClassAtual As String
Public ClassAnterior As String

