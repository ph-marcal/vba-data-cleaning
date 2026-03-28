Attribute VB_Name = "PadronizarColunasAtivas_Senior"
Sub PadronizarColunaAtiva_Senior()
    Dim ws As Worksheet
    Dim rng As Range
    Dim dados As Variant
    Dim i As Long, ultimaLinha As Long, colAtiva As Long
    Dim tempoInicio As Double
    
    ' Configurações de Performance
    On Error GoTo ErroHandler
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .Cursor = xlWait
    End With
    
    tempoInicio = Timer
    Set ws = ActiveSheet
    colAtiva = ActiveCell.Column
    ultimaLinha = ws.Cells(ws.Rows.Count, colAtiva).End(xlUp).Row
    
    If ultimaLinha < 2 Then
        MsgBox "A coluna selecionada está vazia ou contém apenas cabeçalho.", vbExclamation
        GoTo Finalizar
    End If
    
    ' Define o intervalo e carrega o Array
    Set rng = ws.Range(ws.Cells(2, colAtiva), ws.Cells(ultimaLinha, colAtiva))
    
    ' Caso especial: Apenas 1 linha de dados
    If rng.Cells.Count = 1 Then
        dados = ReDimPreserveInput(rng.Value)
    Else
        dados = rng.Value
    End If
    
    ' Processamento em Memória
    For i = 1 To UBound(dados, 1)
        If Not IsError(dados(i, 1)) Then
            If Len(Trim(dados(i, 1))) > 0 Then
                ' 1. UCase + Trim (Removendo espaços duplos internos)
                dados(i, 1) = UCase(Application.Trim(dados(i, 1)))
                
                ' 2. Opcional: Remover acentos (crucial para auditoria/PROCV)
                dados(i, 1) = RemoverAcentos(CStr(dados(i, 1)))
            End If
        End If
    Next i
    
    ' Devolve os dados para a planilha
    rng.Value = dados
    
    ' Log de conclusão
    With Application
        .ScreenUpdating = True
        MsgBox "Processamento concluído em " & Format(Timer - tempoInicio, "0.00") & " segundos." & vbCrLf & _
               "Coluna: " & Split(ws.Cells(1, colAtiva).Address, "$")(1) & " padronizada.", vbInformation
    End With

Finalizar:
    ' Restaura as configurações do Excel
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .Cursor = xlDefault
    End With
    Exit Sub

ErroHandler:
    MsgBox "Erro inesperado: " & Err.Description, vbCritical, "Erro de Auditoria"
    Resume Finalizar
End Sub

' Função auxiliar para garantir que dados de 1 única célula sejam tratados como Array
Private Function ReDimPreserveInput(v As Variant) As Variant
    Dim arr(1 To 1, 1 To 1) As Variant
    arr(1, 1) = v
    ReDimPreserveInput = arr
End Function

' Função Essencial para Auditoria Contábil (Padronização de Strings)
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

