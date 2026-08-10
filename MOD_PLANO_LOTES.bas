Attribute VB_Name = "BatimentoLoteNivelINV"
Option Explicit

' Planeja transferencias entre lotes do mesmo material.
' Regra de negocio preservada: uma falta so pode usar sobras de linhas posteriores.
' O planejamento inteiro falha se qualquer falta nao puder ser coberta integralmente.

Private Const PRIMEIRA_LINHA_DADOS As Long = 6
Private Const COL_MATERIAL As String = "S"
Private Const COL_LOTE As String = "U"
Private Const COL_SAP_LIVRE As String = "V"
Private Const COL_DLX_LIVRE As String = "Y"
Private Const EPSILON As Double = 0.000001

' Retorna uma matriz com:
' material | lote origem | lote destino | quantidade | linha origem | linha destino
' Retorna Empty quando nao ha movimentos necessarios.
Public Sub GerarPreviaBatimentoInventario()
    GerarPreviaBatimentoINV ws_base
End Sub

Public Function PlanejarBatimentoINV(ByVal wsBase As Worksheet) As Variant
    Dim ultimaLinha As Long
    Dim dados As Variant
    Dim saldos() As Double
    Dim materiais() As String
    Dim lotes() As String
    Dim movimentos As Collection
    Dim i As Long, j As Long
    Dim faltaRestante As Double, sobraDisponivel As Double, quantidade As Double

    If wsBase Is Nothing Then Err.Raise vbObjectError + 1000, , "A planilha base nao foi informada."

    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, COL_SAP_LIVRE).End(xlUp).Row
    If ultimaLinha < PRIMEIRA_LINHA_DADOS Then
        PlanejarBatimentoINV = Empty
        Exit Function
    End If

    dados = wsBase.Range(COL_MATERIAL & PRIMEIRA_LINHA_DADOS & ":" & COL_DLX_LIVRE & ultimaLinha).Value2
    ReDim saldos(1 To UBound(dados, 1))
    ReDim materiais(1 To UBound(dados, 1))
    ReDim lotes(1 To UBound(dados, 1))

    For i = 1 To UBound(dados, 1)
        materiais(i) = CodigoTexto(dados(i, 1), "material", i + PRIMEIRA_LINHA_DADOS - 1)
        lotes(i) = CodigoTexto(dados(i, 3), "lote", i + PRIMEIRA_LINHA_DADOS - 1)
        saldos(i) = Quantidade(dados(i, 4), "SAP Livre", i + PRIMEIRA_LINHA_DADOS - 1) _
                  - Quantidade(dados(i, 7), "DLX Livre", i + PRIMEIRA_LINHA_DADOS - 1)
    Next i

    Set movimentos = New Collection

    For i = 1 To UBound(saldos)
        If saldos(i) < -EPSILON Then
            faltaRestante = -saldos(i)

            For j = i + 1 To UBound(saldos)
                If faltaRestante <= EPSILON Then Exit For

                ' Transferir entre materiais diferentes seria um lancamento invalido.
                If materiais(j) = materiais(i) And saldos(j) > EPSILON Then
                    sobraDisponivel = saldos(j)
                    quantidade = MenorValor(faltaRestante, sobraDisponivel)

                    movimentos.Add Array(materiais(i), lotes(j), lotes(i), quantidade, _
                                          j + PRIMEIRA_LINHA_DADOS - 1, i + PRIMEIRA_LINHA_DADOS - 1)
                    saldos(j) = saldos(j) - quantidade
                    saldos(i) = saldos(i) + quantidade
                    faltaRestante = faltaRestante - quantidade
                End If
            Next j

            ' Nenhum movimento e escrito na planilha ou no SAP antes desta validacao terminar.
            If faltaRestante > EPSILON Then
                Err.Raise vbObjectError + 1001, , _
                    "A falta do material " & materiais(i) & " no lote " & lotes(i) _
                    & " (linha " & i + PRIMEIRA_LINHA_DADOS - 1 & ") nao pode ser coberta " _
                    & "integralmente por lotes posteriores do mesmo material."
            End If
        End If
    Next i

    PlanejarBatimentoINV = MatrizDeMovimentos(movimentos)
End Function

' Gera uma previa auditavel em AJ:AO. Nao envia nenhum lancamento ao SAP.
Public Sub GerarPreviaBatimentoINV(ByVal wsBase As Worksheet)
    Dim plano As Variant
    Dim ultimaLinhaSaida As Long

    On Error GoTo Falha
    plano = PlanejarBatimentoINV(wsBase)

    ultimaLinhaSaida = wsBase.Cells(wsBase.Rows.Count, "AJ").End(xlUp).Row
    If ultimaLinhaSaida < PRIMEIRA_LINHA_DADOS Then ultimaLinhaSaida = PRIMEIRA_LINHA_DADOS
    wsBase.Range("AJ" & PRIMEIRA_LINHA_DADOS & ":AO" & ultimaLinhaSaida).ClearContents
    wsBase.Range("AJ5").Resize(1, 6).Value = CabecalhoPrevia()

    If IsEmpty(plano) Then
        MsgBox "Nao ha movimentos a gerar.", vbInformation
    Else
        wsBase.Range("AJ" & PRIMEIRA_LINHA_DADOS).Resize(UBound(plano, 1), UBound(plano, 2)).Value = plano
        MsgBox UBound(plano, 1) & " movimento(s) planejado(s). Revise a previa antes do envio ao SAP.", vbInformation
    End If
    Exit Sub

Falha:
    MsgBox Err.Description, vbExclamation, "Batimento INV"
End Sub

Private Function MatrizDeMovimentos(ByVal movimentos As Collection) As Variant
    Dim resultado() As Variant
    Dim item As Variant
    Dim linha As Long, coluna As Long

    If movimentos.Count = 0 Then
        MatrizDeMovimentos = Empty
        Exit Function
    End If

    ReDim resultado(1 To movimentos.Count, 1 To 6)
    For linha = 1 To movimentos.Count
        item = movimentos(linha)
        For coluna = 1 To 6
            resultado(linha, coluna) = item(coluna - 1)
        Next coluna
    Next linha
    MatrizDeMovimentos = resultado
End Function

Private Function Quantidade(ByVal valor As Variant, ByVal campo As String, ByVal linha As Long) As Double
    If IsError(valor) Then
        Err.Raise vbObjectError + 1002, , campo & " invalido na linha " & linha & "."
    End If
    If Len(Trim$(CStr(valor))) = 0 Or Not IsNumeric(valor) Then Err.Raise vbObjectError + 1002, , campo & " invalido na linha " & linha & "."
    If CDbl(valor) < 0 Then Err.Raise vbObjectError + 1003, , campo & " nao pode ser negativo na linha " & linha & "."
    Quantidade = CDbl(valor)
End Function

Private Function CodigoTexto(ByVal valor As Variant, ByVal campo As String, ByVal linha As Long) As String
    If IsError(valor) Then
        Err.Raise vbObjectError + 1004, , campo & " vazio na linha " & linha & "."
    End If
    If Len(Trim$(CStr(valor))) = 0 Then Err.Raise vbObjectError + 1004, , campo & " vazio na linha " & linha & "."
    CodigoTexto = Trim$(CStr(valor))
End Function

Private Function CabecalhoPrevia() As Variant
    Dim cabecalho(1 To 1, 1 To 6) As Variant

    cabecalho(1, 1) = "Material"
    cabecalho(1, 2) = "Lote origem"
    cabecalho(1, 3) = "Lote destino"
    cabecalho(1, 4) = "Quantidade"
    cabecalho(1, 5) = "Linha origem"
    cabecalho(1, 6) = "Linha destino"
    CabecalhoPrevia = cabecalho
End Function

Private Function MenorValor(ByVal primeiro As Double, ByVal segundo As Double) As Double
    If primeiro < segundo Then
        MenorValor = primeiro
    Else
        MenorValor = segundo
    End If
End Function

