Attribute VB_Name = "MOD_LINK"
Sub AbrirLink_APP_LOTES()

    Dim link As String
    
    link = "https://apps.powerapps.com/play/e/default-42cc3295-cd0e-449c-b98e-5ce5b560c1d3/a/79f3f332-af17-4ca2-b42c-71a9be445d68?tenantId=42cc3295-cd0e-449c-b98e-5ce5b560c1d3&sourcetime=1732884308969" ' Substitua pelo link desejado
    ActiveWorkbook.FollowHyperlink Address:=link
    
End Sub

Sub AbrirLink_APP_CHAMADO()

    Dim link As String
    
    link = "https://apps.powerapps.com/play/e/default-42cc3295-cd0e-449c-b98e-5ce5b560c1d3/a/299ebecf-1e05-43e2-89f9-dd422b7a51d3?tenantId=42cc3295-cd0e-449c-b98e-5ce5b560c1d3&sourcetime=1724676402902" ' Substitua pelo link desejado
    ActiveWorkbook.FollowHyperlink Address:=link
    
End Sub


