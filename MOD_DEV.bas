Attribute VB_Name = "MOD_DEV"
Sub ExibirTudo()

Windows("Batimento Operacional.xlsm").Activate

ws_config.Visible = xlSheetVisible
ws_Arquivos.Visible = xlSheetVisible
ws_base.Visible = xlSheetVisible
ws_dados.Visible = xlSheetVisible


End Sub


Sub Ocultar()

Windows("Batimento Operacional.xlsm").Activate

ws_config.Visible = xlSheetVeryHidden
ws_Arquivos.Visible = xlSheetVeryHidden
ws_base.Visible = xlSheetVeryHidden
ws_dados.Visible = xlSheetVeryHidden


End Sub

Sub MemoriaExcel()
    ' Desativa o modo de cópia/corte
    Application.CutCopyMode = False
    
    ' Força o Excel a recalcular as fórmulas
    Application.Calculate
    
    ' Limpa a área de transferência
    Application.CutCopyMode = False
End Sub
