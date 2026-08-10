Attribute VB_Name = "MOD_CONECT_SHARE"
Sub share_dados_batimento()


Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False

On Error GoTo ERROR



    Dim rs      As ADODB.Recordset
    Dim conn    As ADODB.Connection
    Dim SQL     As String
    Dim Tabela  As ListObject
    Dim linha   As ListRow
    Dim share   As String
    Dim lista_share As String
  
 
 
    Set conn = New ADODB.Connection  ' criando objeto
    Set rs = New ADODB.Recordset
  
ws_dados.Select

With conn
        
    .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes;DATABASE=" & ws_dados.Range("xfd1").Value & ";LIST=" & ws_dados.Range("xfd2").Value & ";"
    .Open
    
    End With
    
    Set Tabela = ws_dados.ListObjects(1)
    
    
    SQL = "Select * FROM [" & ws_dados.ListObjects(1).Name & "]"
    rs.Open SQL, conn, adOpenDynamic, adLockOptimistic
    
    For Each linha In Tabela.ListRows
    
        If Application.WorksheetFunction.CountA(linha.Range) = linha.Range.Columns.Count Then ' INSERIR APENAS LINHAS PREENCHIDAS
    
        rs.AddNew
        rs.Fields("GPID") = linha.Range(, 1).Value
        rs.Fields("DATA") = linha.Range(, 2).Value
        rs.Fields("HORA_INICIO") = linha.Range(, 3).Value
        rs.Fields("HORA_FIM") = linha.Range(, 4).Value
        rs.Fields("MODULO") = linha.Range(, 5).Value
        rs.Fields("LINHAS_REALIZADAS") = linha.Range(, 6).Value
        rs.Fields("LINHAS_POSITIVAS") = linha.Range(, 7).Value
        rs.Fields("LINHAS NEGATIVAS") = linha.Range(, 8).Value
        rs.Fields("TOTAL") = linha.Range(, 9).Value
       
        rs.Update
    End If
       
    
    
    Next linha
    
    

rs.Close ' fechar o recording set
conn.Close
Set rs = Nothing
Set conn = Nothing




Exit Sub

ERROR:

MsgBox "Transmissão de dados falhou , favor abrir um ticket no Aplicativo WH S&D Suport", vbCritical, "Erro na transmissao de dados"


End Sub

