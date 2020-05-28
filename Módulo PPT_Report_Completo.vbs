Option Compare Text

Sub Report_Completo_PPT()
    
    'Cria uma nova aplicação do Powerpoint:
    Set appPPT = New PowerPoint.Application
    
    'Torna essa aplicação visível:
    appPPT.Visible = True
    'maximize PowerPoint window:
    'appPPT.WindowState = ppWindowMaximized
    
    'Cria uma nova apresentação no Powerpoint:
    Set prsntPPT = appPPT.Presentations.Add
        prsntPPT.PageSetup.SlideSize = ppSlideSizeOnScreen
        
    SlideCountPPT = 1
    Total_RowCountPPT = 0
    RowCountPPT = 2
    
    Call OrganizaTabelaPPT
    Call NovoSlide
    Call PreencheTabelaPPT

'    Do Until Total_RowCountPPT >= Tabela.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count
'        Call PreencheTabelaPPT
'    Loop
    
    'Define o caminho até o arquivo, por definição o mesmo do excel
    PathPPT = ThisWorkbook.Path
    
    'Define o nome do ppt
    NamePPT = PathPPT & "/" & Format(Now(), "yyyymmdd") & " - " & "Report Completo " & Planilha.Name & ".pptx"
    
    'Salva arquivoPPT
    prsntPPT.SaveAs NamePPT

End Sub

Sub CriaTabelaPPT()

'Cria o Header da tabela, formata primeiro o tamanho das colunas, depois preenche cada uma delas com o texto formatado da forma desejada

'Report LUDO
'#1 Projeto
'#2 Status Anterior
'#3 Status Atual
'#4 Fase Atual
'#5 Data Fim (prevista)
'#6 Data Fim (replanej)
'#7 Status Report
'#8 Gestor

Set shapePPT = slidePPT.Shapes.AddTable(NumRows:=1, NumColumns:=8, Left:=7.69, Top:=76, Width:=685.44) 'Define o Tamanho e Dimensões
    'Width Máximo: 707
    With shapePPT.Table
        .Columns(1).Width = 130
        .Columns(2).Width = 53
        .Columns(3).Width = 53
        .Columns(4).Width = 78
        .Columns(5).Width = 63
        .Columns(6).Width = 63
        .Columns(7).Width = 189
        .Columns(8).Width = 78
        With .Cell(1, 1).Shape
            With .TextFrame
                .TextRange.Text = "Projeto"
                .TextRange.ParagraphFormat.Alignment = ppAlignCenter
                .TextRange.Font.Size = 11
                .VerticalAnchor = msoAnchorMiddle
            End With
        End With
        
        With .Cell(1, 2).Shape
            With .TextFrame
                .TextRange.Text = "Status" & Chr(10) & "Anterior"
                .TextRange.ParagraphFormat.Alignment = ppAlignCenter
                .TextRange.Font.Size = 11
                .VerticalAnchor = msoAnchorMiddle
            End With
        End With
        
        With .Cell(1, 3).Shape
            With .TextFrame
                .TextRange.Text = "Status" & Chr(10) & "Atual"
                .TextRange.ParagraphFormat.Alignment = ppAlignCenter
                .TextRange.Font.Size = 11
                .VerticalAnchor = msoAnchorMiddle
            End With
        End With
        
        With .Cell(1, 4).Shape
            With .TextFrame
                .TextRange.Text = "Fase" & Chr(10) & "Atual"
                .TextRange.ParagraphFormat.Alignment = ppAlignCenter
                .VerticalAnchor = msoAnchorMiddle
                .TextRange.Font.Size = 11
            End With
        End With
        
        With .Cell(1, 5).Shape
            With .TextFrame
                .TextRange.Text = "Data Fim" & Chr(10) & "(prevista)"
                .TextRange.ParagraphFormat.Alignment = ppAlignCenter
                .TextRange.Font.Size = 11
                .VerticalAnchor = msoAnchorMiddle
            End With
        End With
        
        With .Cell(1, 6).Shape
            With .TextFrame
                .TextRange.Text = "Data Fim" & Chr(10) & "(replanej)"
                .TextRange.ParagraphFormat.Alignment = ppAlignCenter
                .TextRange.Font.Size = 11
                .VerticalAnchor = msoAnchorMiddle
            End With
        End With
        
        With .Cell(1, 7).Shape
            With .TextFrame
                .TextRange.Text = "Status Report"
                .TextRange.ParagraphFormat.Alignment = ppAlignLeft
                .TextRange.Font.Size = 11
                .VerticalAnchor = msoAnchorMiddle
            End With
        End With
        
        With .Cell(1, 8).Shape
            With .TextFrame
                .TextRange.Text = "GP"
                .TextRange.ParagraphFormat.Alignment = ppAlignCenter
                .TextRange.Font.Size = 11
                .VerticalAnchor = msoAnchorMiddle
            End With
        End With
    End With
 'Aplica a formatação bonitinha... esse número eu achei aqui https://msdn.microsoft.com/en-us/library/office/hh273476(v=office.14).aspx
 'Ele representa a entrada do estilo de tabela a ser usado
    shapePPT.Table.ApplyStyle "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}", False
End Sub

Sub OrganizaTabelaPPT()

    'Organiza a Tabela da forma desejada
    Tabela.Sort.SortFields.Clear
    Tabela.Sort.SortFields.Add Key:=Range(Tabela & "[" & Class1 & "]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        
    If Class2 <> "" Then _
        Tabela.Sort.SortFields.Add Key:=Range(Tabela & "[" & Class2 & "]"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        
    If Class3 <> "" Then _
        Tabela.Sort.SortFields.Add Key:=Range(Tabela & "[" & Class3 & "]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        
    With Tabela.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
If Class1 = "ID'#" Then Class1 = "ID#"
If Class2 = "ID'#" Then Class2 = "ID#"
If Class3 = "ID'#" Then Class3 = "ID#"
    
    'Filtra a tabela da forma desejada

'    If Filter1 <> "" Then
'        Tabela.AutoFilter.ShowAllData
'        Tabela.DataBodyRange.AutoFilter Tabela.ListColumns(Filter1).Index, FilterAs1
'    End If
'    If Filter2 <> "" Then
'        Tabela.DataBodyRange.AutoFilter Tabela.ListColumns(Filter1).Index, FilterAs1
'    End If
End Sub

Sub PreencheTabelaPPT()

For Each VisibleRow In Tabela.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows

'Armazena o Class1 que será usado nessa iteração
ClassAtual = Intersect(VisibleRow, Tabela.ListColumns(Class1).DataBodyRange).Value
    
    
    'Fluxos de exceção para ID e Datas, para que ao mudar esses IDs ou Datas, não seja gerado um novo slide
    If ((Class1 <> "ID1#" And Class1 <> "ID#") And _
    Class1 <> "Data da Recomendação" And _
    Class1 <> "Data Início" And Class1 <> "Data fim - planejada" And _
    Class1 <> "Data fim-Replanejada" And _
    Class1 <> "Data Término" And Class1 <> "Ultima data planejada" And _
    Class1 <> "Aging") And _
    (ClassAnterior <> ClassAtual And ClassAnterior <> "") Then GoTo MudouClass1 'Essa linha é o real motivo do If. Se mudar o Class1, fora das exceções, cria um novo Slide
    
tryagain:

        'Define o que vai em cada coluna, limitado aos primeiros 255 caractéres, o -3 é o offset da primeira linha de dados em relação ao inicio da planilha
        Coluna1 = Left(Tabela.DataBodyRange.Cells(VisibleRow.Row - 3, 1), 255) 'Projeto
        Coluna2 = Left(Tabela.DataBodyRange.Cells(VisibleRow.Row - 3, 60), 255) 'Status Anterior
        Coluna3 = Left(Tabela.DataBodyRange.Cells(VisibleRow.Row - 3, 61), 255) 'Status Atual
        Coluna4 = Left(Tabela.DataBodyRange.Cells(VisibleRow.Row - 3, 11), 255) 'Fase MGP
        Coluna5 = Left(Tabela.DataBodyRange.Cells(VisibleRow.Row - 3, 66), 255) 'Data Fim (Prevista)
        Coluna6 = Left(Tabela.DataBodyRange.Cells(VisibleRow.Row - 3, 69), 255) 'Data Fim (Replanej)
        Coluna7 = Left(Tabela.DataBodyRange.Cells(VisibleRow.Row - 3, 62), 700) 'Status Report
        Coluna8 = Left(Tabela.DataBodyRange.Cells(VisibleRow.Row - 3, 55), 255) 'Gestor
        
    
        'Preenche a tabela, linha a linha, até ter um número igual ao da Tabela
        shapePPT.Table.Rows.Add
        shapePPT.Table.Cell(RowCountPPT, 1).Shape.TextFrame.TextRange.Text = Coluna1
        shapePPT.Table.Cell(RowCountPPT, 2).Shape.TextFrame.TextRange.Text = Coluna2
        shapePPT.Table.Cell(RowCountPPT, 3).Shape.TextFrame.TextRange.Text = Coluna3
        shapePPT.Table.Cell(RowCountPPT, 4).Shape.TextFrame.TextRange.Text = Coluna4
        shapePPT.Table.Cell(RowCountPPT, 5).Shape.TextFrame.TextRange.Text = Format(Coluna5, "dd/mm/yy")
        shapePPT.Table.Cell(RowCountPPT, 6).Shape.TextFrame.TextRange.Text = Format(Coluna6, "dd/mm/yy")
        shapePPT.Table.Cell(RowCountPPT, 7).Shape.TextFrame.TextRange.Text = Coluna7
        shapePPT.Table.Cell(RowCountPPT, 8).Shape.TextFrame.TextRange.Text = Coluna8
        
            
        'Formatar Campo de Status
        FormataStatus (2)
        FormataStatus (3)

        If shapePPT.Height >= 440 Then GoTo TabelaGrande 'caso a tabela venha a ficar maior que o slide, cria 1 novo
               
        RowCountPPT = RowCountPPT + 1
                           
        'Se tudo der certo, finaliza a sub
        GoTo Beleza:

MudouClass1:
        RowCountPPT = 2
        Call CriaCabeçalhoPPT
        Call NovoSlide
        GoTo tryagain
    
    
TabelaGrande:
        shapePPT.Table.Rows(RowCountPPT).Delete
        RowCountPPT = 2
        Call CriaCabeçalhoPPT
        Call NovoSlide
        GoTo tryagain
    
Beleza:

        'Armazena qual o Class1 que foi usado no loop anterior
        ClassAnterior = Intersect(VisibleRow.EntireRow, Tabela.ListColumns(Class1).DataBodyRange).Value
        Next VisibleRow
        Call CriaCabeçalhoPPT

        'Limpa o ClassAtual e Anterior para não comprometer novos loops
        ClassAnterior = ""
        ClassAtual = ""
End Sub

Sub FormataStatus(NumColuna As Integer)

With shapePPT.Table.Cell(RowCountPPT, NumColuna).Shape.TextFrame.TextRange
        Select Case .Text
            Case Is = "Em Acompanhamento"
                With .Font
                    .Size = 24
                    .Name = "Wingdings"
                    .Color.RGB = RGB(0, 176, 80)
                End With
                .Text = "Ý"
                
            Case Is = "Em ação imediata"
                With .Font
                    .Size = 24
                    .Name = "Wingdings"
                    .Color.RGB = RGB(255, 0, 0)
                End With
                .Text = "Þ"
            
            Case Is = "Não Iniciado"
                With .Font
                    .Size = 24
                    .Name = "Webdings"
                    .Color.RGB = RGB(0, 0, 0)
                End With
                .Text = "y"
    
            Case Is = "Finalizado"
                With .Font
                    .Size = 24
                    .Name = "Wingdings 2"
                    .Color.RGB = RGB(0, 112, 192)
                End With
                .Text = "P"
    
            Case Is = "Em Alerta"
                With .Font
                    .Size = 24
                    .Name = "Wingdings"
                    .Color.RGB = RGB(255, 204, 0)
                End With
                .Text = "Ü"
    
            Case Is = "Suspenso"
                With .Font
                    .Size = 24
                    .Name = "webdings"
                    .Color.RGB = RGB(0, 0, 0)
                End With
                .Text = ";"
                
            Case Is = "Cancelado"
                With .Font
                    .Size = 24
                    .Name = "Wingdings 2"
                    .Color.RGB = RGB(255, 0, 0)
                End With
                .Text = "U"
            
            Case Else
                'End
        End Select
    End With

End Sub

Sub NovoSlide()
    'Cria um Slide
    Set slidePPT = prsntPPT.Slides.Add(SlideCountPPT, ppLayoutBlank)
    SlideCountPPT = SlideCountPPT + 1
    Call CriaTabelaPPT
End Sub

Sub CriaCabeçalhoPPT()

    Set headerPPT = slidePPT.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=8.31, Top:=0, Width:=638.46, Height:=70.31)
        
        
      'Exceção quando o Class1 for data ou ID
        If ((Class1 <> "ID1#" And Class1 <> "ID#") And _
            Class1 <> "Data da Recomendação" And _
            Class1 <> "Data Início" And Class1 <> "Data fim - planejada" And _
            Class1 <> "Data fim-Replanejada" And _
            Class1 <> "Data Término" And Class1 <> "Ultima data planejada" And _
            Class1 <> "Aging") Then
           'Caso sem ser exceção
            headerPPT.TextFrame.TextRange.Text = "Comitê " & Planilha.Name & Chr(10) & Class1 & ": " & ClassAnterior
        Else 'Exceção
            headerPPT.TextFrame.TextRange.Text = "Ambiente " & Planilha.Name & Chr(10) & Class1
        End If
        With headerPPT.TextFrame.TextRange.Lines(1).Font
            .Size = 24
            .Color.RGB = RGB(89, 89, 89)
        End With
        With headerPPT.TextFrame.TextRange.Lines(2).Font
            .Size = 20
            .Color.RGB = RGB(89, 89, 89)
        End With
        
    Set headerlinePPT = slidePPT.Shapes.AddLine(BeginX:=8.31, BeginY:=65.85, EndX:=510, EndY:=65.85)
        headerlinePPT.Line.ForeColor.RGB = RGB(237, 125, 49)
        
End Sub
