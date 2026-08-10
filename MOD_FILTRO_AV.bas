Attribute VB_Name = "MOD_FILTRO_AV"
Sub Filtro_Avancado()

'Filtro avançado

ws_base.Range("U5").CurrentRegion.Clear

ws_base.Range("A5").CurrentRegion.AdvancedFilter xlFilterCopy, ws_base.Range("A1").CurrentRegion, ws_base.Range("R5")

End Sub

Sub LimparMemoriaExcel()

    ' Desativa o modo de cópia/corte
    Application.CutCopyMode = False

    ' Força o Excel a recalcular as fórmulas
    Application.Calculate

    ' Limpa a área de transferência
    Application.CutCopyMode = False ' Esta linha já foi utilizada, pode ser removida

End Sub

