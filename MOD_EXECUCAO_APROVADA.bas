Attribute VB_Name = "MOD_EXECUCAO_APROVADA"
Option Explicit

' Esta camada impede que as macros de SAP sejam chamadas sem uma previa.
' As macros legadas continuam isoladas em legacy/ ate a migracao completa.

Private Const LINHA_CABECALHO_PREVIA As Long = 5
Private Const PRIMEIRA_LINHA_PREVIA As Long = 6

Public Sub ExecutarBatimentoInventarioAprovado()
    If Not ConfirmarExecucaoDaPrevia("Batimento de Inventario") Then Exit Sub

    ' Compatibilidade temporaria: a rotina legada ainda contem a navegacao SAP.
    ' A proxima etapa substitui essa chamada pelo envio direto da matriz AJ:AO.
    batimento_lote_nivel_INV
End Sub

Public Sub ExecutarBatimentoDiarioAprovado()
    If Not ConfirmarExecucaoDaPrevia("Batimento Diario") Then Exit Sub
    batimento_lote_nivel_diario
End Sub

Public Sub ExecutarBatimentoBloqueioAprovado()
    If Not ConfirmarExecucaoDaPrevia("Batimento de Bloqueio") Then Exit Sub
    sb_Batimento_Status_BLOQUEIO
End Sub

Public Sub ExecutarBatimentoDesbloqueioAprovado()
    If Not ConfirmarExecucaoDaPrevia("Batimento de Desbloqueio") Then Exit Sub
    sb_Batimento_Status_desbloqueio
End Sub

Private Function ConfirmarExecucaoDaPrevia(ByVal titulo As String) As Boolean
    Dim ultimaLinha As Long
    Dim quantidadeMovimentos As Long

    If ws_base.Range("AJ" & LINHA_CABECALHO_PREVIA).Value <> "Material" Then
        MsgBox "Gere uma previa antes de executar " & titulo & ".", vbExclamation, titulo
        Exit Function
    End If

    ultimaLinha = ws_base.Cells(ws_base.Rows.Count, "AJ").End(xlUp).Row
    If ultimaLinha < PRIMEIRA_LINHA_PREVIA Then
        MsgBox "A previa nao possui movimentos para executar.", vbInformation, titulo
        Exit Function
    End If

    quantidadeMovimentos = Application.WorksheetFunction.CountA( _
        ws_base.Range("AJ" & PRIMEIRA_LINHA_PREVIA & ":AJ" & ultimaLinha))
    If quantidadeMovimentos = 0 Then
        MsgBox "A previa nao possui movimentos para executar.", vbInformation, titulo
        Exit Function
    End If

    ConfirmarExecucaoDaPrevia = (MsgBox( _
        "A previa contem " & quantidadeMovimentos & " movimento(s)." & vbCrLf & vbCrLf _
        & "Confirma o envio ao SAP?", vbQuestion + vbYesNo + vbDefaultButton2, titulo) = vbYes)
End Function

