Attribute VB_Name = "MOD_PLANO_STATUS"
Option Explicit

Private Const LINHA_INICIAL_STATUS As Long = 6

Public Sub GerarPreviaStatusBloqueio()
    GerarPreviaStatus True
End Sub

Public Sub GerarPreviaStatusDesbloqueio()
    GerarPreviaStatus False
End Sub

' AC e o saldo de bloqueio. Positivo bloqueia (344); negativo desbloqueia (343).
Private Sub GerarPreviaStatus(ByVal bloquear As Boolean)
    Dim ultimaLinha As Long, i As Long, qtd As Double, diferenca As Double
    Dim saida() As Variant, total As Long, dados As Variant

    On Error GoTo Falha
    ultimaLinha = ws_base.Cells(ws_base.Rows.Count, "S").End(xlUp).Row
    If ultimaLinha < LINHA_INICIAL_STATUS Then Exit Sub
    dados = ws_base.Range("S" & LINHA_INICIAL_STATUS & ":AC" & ultimaLinha).Value2
    ReDim saida(1 To UBound(dados, 1), 1 To 6)

    For i = 1 To UBound(dados, 1)
        If Not IsNumeric(dados(i, 11)) Then Err.Raise vbObjectError + 1200, , "Saldo de bloqueio invalido na linha " & i + 5 & "."
        diferenca = CDbl(dados(i, 11)) ' AC: saldo bloqueado calculado
        If (bloquear And diferenca > 0) Or (Not bloquear And diferenca < 0) Then
            If bloquear Then
                qtd = MenorStatus(diferenca, NumeroStatus(dados(i, 4), "SAP Livre", i + 5))
            Else
                qtd = MenorStatus(-diferenca, NumeroStatus(dados(i, 5), "SAP Bloqueado", i + 5))
            End If
            If qtd > 0 Then
                total = total + 1
                saida(total, 1) = CodigoStatus(dados(i, 1), "material", i + 5)
                saida(total, 2) = CodigoStatus(dados(i, 3), "lote", i + 5)
                saida(total, 3) = IIf(bloquear, "344", "343")
                saida(total, 4) = qtd
                saida(total, 5) = IIf(bloquear, "Livre para Bloqueado", "Bloqueado para Livre")
                saida(total, 6) = i + 5
            End If
        End If
    Next i

    ws_base.Range("AJ6:AO" & ws_base.Cells(ws_base.Rows.Count, "AJ").End(xlUp).Row).ClearContents
    ws_base.Range("AI4").Value = IIf(bloquear, "BLOQUEIO", "DESBLOQUEIO")
    ws_base.Range("AJ5:AO5").Value = CabecalhoStatus()
    If total = 0 Then
        MsgBox "Nao ha movimentos de status a gerar.", vbInformation
    Else
        ws_base.Range("AJ6").Resize(total, 6).Value = RecortarStatus(saida, total)
        MsgBox total & " movimento(s) planejado(s).", vbInformation
    End If
    Exit Sub
Falha:
    MsgBox Err.Description, vbExclamation, "Previa de Status"
End Sub

Private Function RecortarStatus(ByRef origem() As Variant, ByVal total As Long) As Variant
    Dim destino() As Variant, i As Long, j As Long
    ReDim destino(1 To total, 1 To 6)
    For i = 1 To total: For j = 1 To 6: destino(i, j) = origem(i, j): Next j, i
    RecortarStatus = destino
End Function

Private Function CabecalhoStatus() As Variant
    Dim h(1 To 1, 1 To 6) As Variant
    h(1, 1) = "Material": h(1, 2) = "Lote": h(1, 3) = "Movimento": h(1, 4) = "Quantidade": h(1, 5) = "Direcao": h(1, 6) = "Linha"
    CabecalhoStatus = h
End Function

Private Function NumeroStatus(ByVal valor As Variant, ByVal campo As String, ByVal linha As Long) As Double
    If IsError(valor) Or Not IsNumeric(valor) Then Err.Raise vbObjectError + 1201, , campo & " invalido na linha " & linha & "."
    If CDbl(valor) < 0 Then Err.Raise vbObjectError + 1202, , campo & " negativo na linha " & linha & "."
    NumeroStatus = CDbl(valor)
End Function

Private Function CodigoStatus(ByVal valor As Variant, ByVal campo As String, ByVal linha As Long) As String
    CodigoStatus = Trim$(CStr(valor))
    If Len(CodigoStatus) = 0 Then Err.Raise vbObjectError + 1203, , campo & " vazio na linha " & linha & "."
End Function

Private Function MenorStatus(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MenorStatus = a Else MenorStatus = b
End Function

