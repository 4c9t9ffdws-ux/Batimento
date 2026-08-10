Attribute VB_Name = "MOD_SUPORTE"
Option Explicit

Private Const URL_APP_LOTES As String = "https://apps.powerapps.com/play/e/default-42cc3295-cd0e-449c-b98e-5ce5b560c1d3/a/79f3f332-af17-4ca2-b42c-71a9be445d68?tenantId=42cc3295-cd0e-449c-b98e-5ce5b560c1d3&sourcetime=1732884308969"
Private Const URL_APP_CHAMADO As String = "https://apps.powerapps.com/play/e/default-42cc3295-cd0e-449c-b98e-5ce5b560c1d3/a/299ebecf-1e05-43e2-89f9-dd422b7a51d3?tenantId=42cc3295-cd0e-449c-b98e-5ce5b560c1d3&sourcetime=1724676402902"

Public Sub AbrirLink_APP_LOTES()
    ThisWorkbook.FollowHyperlink Address:=URL_APP_LOTES
End Sub

Public Sub AbrirLink_APP_CHAMADO()
    ThisWorkbook.FollowHyperlink Address:=URL_APP_CHAMADO
End Sub

Public Sub Filtro_Avancado()
    Dim ultimaLinhaSaida As Long
    Dim ultimaLinhaBase As Long

    ultimaLinhaSaida = ws_base.Cells(ws_base.Rows.Count, "R").End(xlUp).Row
    If ultimaLinhaSaida >= 6 Then ws_base.Range("R6:AG" & ultimaLinhaSaida).ClearContents

    ultimaLinhaBase = ws_base.Cells(ws_base.Rows.Count, "A").End(xlUp).Row
    If ultimaLinhaBase < 6 Then Exit Sub

    ws_base.Range("A5:P" & ultimaLinhaBase).AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=ws_base.Range("A1:P2"), _
        CopyToRange:=ws_base.Range("R5:AG5"), _
        Unique:=False
End Sub

Public Sub LimparMemoriaExcel()
    Application.CutCopyMode = False
    Application.Calculate
End Sub

Public Sub MemoriaExcel()
    LimparMemoriaExcel
End Sub

Public Sub ExibirTudo()
    DefinirAbasTecnicasVisiveis xlSheetVisible
End Sub

Public Sub Ocultar()
    DefinirAbasTecnicasVisiveis xlSheetVeryHidden
End Sub

Private Sub DefinirAbasTecnicasVisiveis(ByVal visibilidade As XlSheetVisibility)
    ws_config.Visible = visibilidade
    ws_Arquivos.Visible = visibilidade
    ws_base.Visible = visibilidade
    ws_dados.Visible = visibilidade
End Sub

