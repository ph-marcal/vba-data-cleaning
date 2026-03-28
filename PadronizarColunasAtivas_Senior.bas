Attribute VB_Name = "PadronizarColunasAtivas_Senior"
Sub PadronizarColunaAtiva_Senior()
    Dim ws As Worksheet
    Dim rng As Range
    Dim dados As Variant
    Dim i As Long, ultimaLinha As Long, colAtiva As Long
    Dim valorLimpo As String
    
    On Error GoTo ErroHandler
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    Set ws = ActiveSheet
    colAtiva = ActiveCell.Column
    ultimaLinha = ws.Cells(ws.Rows.Count, colAtiva).End(xlUp).Row
    
    If ultimaLinha < 2 Then Exit Sub
    
    Set rng = ws.Range(ws.Cells(2, colAtiva), ws.Cells(ultimaLinha, colAtiva))
    
    ' Tratamento para caso de célula única
    If rng.Cells.Count = 1 Then
        ReDim dados(1 To 1, 1 To 1)
        dados(1, 1) = rng.Value
    Else
        dados = rng.Value
    End If
    
    For i = 1 To UBound(dados, 1)
        If Not IsError(dados(i, 1)) Then
            If Len(Trim(dados(i, 1))) > 0 Then
                
                ' 1. Converte para String para evitar erros de tipo
                valorLimpo = CStr(dados(i, 1))
                
                ' 2. TRUQUE DE MESTRE: Substitui o Espaço ASCII 160 (comum em ERPs) por Espaço Comum
                valorLimpo = Replace(valorLimpo, Chr(160), " ")
                
                ' 3. Remove caracteres não imprimíveis (quebras de linha, tabs, etc)
                valorLimpo = Application.WorksheetFunction.Clean(valorLimpo)
                
                ' 4. Aplica o Trim do Excel (que remove espaços duplos internos) e UCase
                valorLimpo = UCase(Application.Trim(valorLimpo))
                
                ' 5. Remove acentos (opcional, mas recomendado para auditoria)
                dados(i, 1) = RemoverAcentos(valorLimpo)
                
            End If
        End If
    Next i
    
    rng.Value = dados
    
    MsgBox "Coluna padronizada com sucesso (incluindo limpeza de espaços invisíveis)!", vbInformation

Finalizar:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    Exit Sub
ErroHandler:
    MsgBox "Erro: " & Err.Description
    Resume Finalizar
End Sub

' Mantenha a função RemoverAcentos abaixo...
Private Function RemoverAcentos(Texto As String) As String
    Dim ComAcentos As String, SemAcentos As String
    Dim i As Integer
    ComAcentos = "ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇÑ"
    SemAcentos = "AAAAAEEEEIIIIOOOOOUUUUCN"
    For i = 1 To Len(ComAcentos)
        Texto = Replace(Texto, Mid(ComAcentos, i, 1), Mid(SemAcentos, i, 1))
    Next i
    RemoverAcentos = Texto
End Function

