Private Sub CommandButton1_Click()
    Call Report_Completo_PPT
    Unload PPT_Report
End Sub

Private Sub Userform_Initialize()
    For Each Value In Tabela.HeaderRowRange
        Me.Classifica1.AddItem Value
        Me.Classifica2.AddItem Value
        Me.Classifica3.AddItem Value
    Next Value
End Sub

Private Sub Classifica1_Change()
    Class1 = Me.Classifica1.Value
    If Class1 = "ID#" Then Class1 = "ID'#"
End Sub

Private Sub Classifica2_Change()
    Class2 = Me.Classifica2.Value
    If Class2 = "ID#" Then Class1 = "ID'#"
End Sub

Private Sub Classifica3_Change()
    Class3 = Me.Classifica3.Value
    If Class3 = "ID#" Then Class1 = "ID'#"
End Sub

