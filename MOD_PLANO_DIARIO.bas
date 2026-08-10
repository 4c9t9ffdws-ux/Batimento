Attribute VB_Name = "MOD_PLANO_DIARIO"
Option Explicit

Private Const PRIMEIRA_LINHA_DIARIO As Long = 6
Private Const EPSILON_DIARIO As Double = 0.000001

Public Sub GerarPreviaBatimentoDiario()
    Dim plano As Variant
    Dim ultimaSaida As Long

    On Error GoTo Falha
    plano = PlanejarBatimentoDiario(ws_base)

    ultimaSaida = ws_base.Cells(ws_base.Rows.Count, "AJ").End(xlUp).Row
    If ultimaSaida < PRIMEIRA_LINHA_DIARIO Then ultimaSaida = PRIMEIRA_LINHA_DIARIO
    ws_base.Range("AJ" & PRIMEIRA_LINHA_DIARIO & ":AO" & ultimaSaida).ClearContents
    ws_base.Range("AI4").Value = "DIARIO"
    ws_base.Range("AJ5").Resize(1, 6).Value = CabecalhoDiario()

    If IsEmpty(plano) Then
        MsgBox "Nao ha movimentos diarios a gerar.", vbInformation, "Batimento Diario"
    Else
        ws_base.Range("AJ" & PRIMEIRA_LINHA_DIARIO).Resize(UBound(plano, 1), UBound(plano, 2)).Value = plano
        MsgBox UBound(plano, 1) & " movimento(s) diario(s) planejado(s).", vbInformation, "Batimento Diario"
    End If
    Exit Sub
Falha:
    MsgBox Err.Description, vbExclamation, "Batimento Diario"
End Sub

' Mantem a prioridade original: procurar sobras da ultima linha para a primeira.
Public Function PlanejarBatimentoDiario(ByVal ws As Worksheet) As Variant
    Dim ultimaLinha As Long, i As Long, j As Long
    Dim dados As Variant, saldo() As Double, material() As String, lote() As String
    Dim falta As Double, quantidade As Double
    Dim movimentos As Collection

    If ws Is Nothing Then Err.Raise vbObjectError + 1100, , "Planilha BASE nao informada."
    ultimaLinha = ws.Cells(ws.Rows.Count, "V").End(xlUp).Row
    If ultimaLinha < PRIMEIRA_LINHA_DIARIO Then Exit Function

    dados = ws.Range("S" & PRIMEIRA_LINHA_DIARIO & ":Y" & ultimaLinha).Value2
    ReDim saldo(1 To UBound(dados, 1))
    ReDim material(1 To UBound(dados, 1))
    ReDim lote(1 To UBound(dados, 1))

    For i = 1 To UBound(dados, 1)
        material(i) = TextoDiario(dados(i, 1), "material", i + PRIMEIRA_LINHA_DIARIO - 1)
        lote(i) = TextoDiario(dados(i, 3), "lote", i + PRIMEIRA_LINHA_DIARIO - 1)
        saldo(i) = NumeroDiario(dados(i, 4), "SAP Livre", i + PRIMEIRA_LINHA_DIARIO - 1) _
                 - NumeroDiario(dados(i, 7), "DLX Livre", i + PRIMEIRA_LINHA_DIARIO - 1)
    Next i

    Set movimentos = New Collection
    For i = 1 To UBound(saldo)
        If saldo(i) < -EPSILON_DIARIO Then
            falta = -saldo(i)
            For j = UBound(saldo) To 1 Step -1
                If falta <= EPSILON_DIARIO Then Exit For
                If j <> i And material(j) = material(i) And saldo(j) > EPSILON_DIARIO Then
                    quantidade = MinimoDiario(falta, saldo(j))
                    movimentos.Add Array(material(i), lote(j), lote(i), quantidade, _
                                         j + PRIMEIRA_LINHA_DIARIO - 1, i + PRIMEIRA_LINHA_DIARIO - 1)
                    saldo(j) = saldo(j) - quantidade
                    saldo(i) = saldo(i) + quantidade
                    falta = falta - quantidade
                End If
            Next j
            If falta > EPSILON_DIARIO Then
                Err.Raise vbObjectError + 1101, , "Falta sem cobertura integral: material " & material(i) _
                    & ", lote " & lote(i) & ", linha " & i + PRIMEIRA_LINHA_DIARIO - 1 & "."
            End If
        End If
    Next i
    PlanejarBatimentoDiario = MatrizDiaria(movimentos)
End Function

Private Function MatrizDiaria(ByVal movimentos As Collection) As Variant
    Dim resultado() As Variant, item As Variant, i As Long, j As Long
    If movimentos.Count = 0 Then Exit Function
    ReDim resultado(1 To movimentos.Count, 1 To 6)
    For i = 1 To movimentos.Count
        item = movimentos(i)
        For j = 1 To 6: resultado(i, j) = item(j - 1): Next j
    Next i
    MatrizDiaria = resultado
End Function

Private Function CabecalhoDiario() As Variant
    Dim itens(1 To 1, 1 To 6) As Variant
    itens(1, 1) = "Material": itens(1, 2) = "Lote origem": itens(1, 3) = "Lote destino"
    itens(1, 4) = "Quantidade": itens(1, 5) = "Linha origem": itens(1, 6) = "Linha destino"
    CabecalhoDiario = itens
End Function

Private Function NumeroDiario(ByVal valor As Variant, ByVal campo As String, ByVal linha As Long) As Double
    If IsError(valor) Then Err.Raise vbObjectError + 1102, , campo & " invalido na linha " & linha & "."
    If Len(Trim$(CStr(valor))) = 0 Or Not IsNumeric(valor) Then Err.Raise vbObjectError + 1102, , campo & " invalido na linha " & linha & "."
    If CDbl(valor) < 0 Then Err.Raise vbObjectError + 1103, , campo & " negativo na linha " & linha & "."
    NumeroDiario = CDbl(valor)
End Function

Private Function TextoDiario(ByVal valor As Variant, ByVal campo As String, ByVal linha As Long) As String
    If IsError(valor) Then Err.Raise vbObjectError + 1104, , campo & " invalido na linha " & linha & "."
    TextoDiario = Trim$(CStr(valor))
    If Len(TextoDiario) = 0 Then Err.Raise vbObjectError + 1104, , campo & " vazio na linha " & linha & "."
End Function

Private Function MinimoDiario(ByVal primeiro As Double, ByVal segundo As Double) As Double
    If primeiro < segundo Then MinimoDiario = primeiro Else MinimoDiario = segundo
End Function

