Attribute VB_Name = "MOD_BASE"
'=============================================================================================================
'variaveis global

Public wsBaseLN            As Integer ' variavel de linha para base
Public resposta            As String ' variavel para armazenar informaçao do tipo de batimento
Public resposta2           As String ' variavel para armazenar o centro
Public resposta3           As String ' variavel para armazenar o tipo de saldo que vai usar
Public LnTBL               As Long   ' variavel para contar as linhas dos saldos
Public lnValidacao         As Long   'variavel para validacao da nossa chave e retirar duplicadas
Public chave               As String ' armazenar chave
Public material            As String ' armazenar material
Public lote                As String 'armazenar lote
Public ln                  As Long 'armazenar linha dos dados filtrados
Public lnarq               As Long 'armazenar linha da aba de arquivos
Public dictDesbloqueio     As Object ' dicionario para armazenar os materias nao duplicados
Public dictBloqueio        As Object ' dicionario para armazenar os materias nao duplicados
Public dictLote            As Object ' dicionario para armazenar os materias nao duplicados
Public linhaarquivos       As Long 'armazenar linha da aba de arquivos
Public qtdSAP              As Long 'armazenar quantidade do SAP
Public qtdNTV              As Long ' armazenar quantidade do Saldo NTV
Public tipZM               As String 'armazenar tipo de ZM
Public V_aux               As Integer ' Auxiliar , armazenar a linha quando filtrar o sku
Public l_Lote              As Integer ' Armazenar quantidade de linhas que foi realizado batimento
Public l_faltante          As Integer ' Armazenar quantidade de linhas que contem valores faltando, sku todo faltando
Public l_Sobrante          As Integer ' Armazenar quantidade de linhas que contem valores sobrante , sku todo sobrando
Public l_bloqueio          As Integer  ' armazenar quantidade bloqueada
Public l_desbloqueio       As Integer ' armazenar quantidade desbloqueada
Public eErros              As Integer ' armazenar linhas com erros
Public validationSaldo     As Boolean ' verificar se o saldo selecionado e correto
Public cCentro             As String ' Armazenar centro
Public CENTRO              As String 'armazenar cod centro



Sub Gerar_base()

Windows("Batimento Operacional.xlsm").Activate
'REIXIBIR ABAS
'=============================================================================================================
ExibirTudo


'=============================================================================================================
'limpar memoria Excel

MemoriaExcel

'=============================================================================================================
'variaveis

 Dim possuiValorDiferenteDeZero As Boolean
 Dim batSobrante As Boolean, batFaltante As Boolean

'=============================================================================================================
'DESABILITAR VISUALIZAÇOES DO EXCEL

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False


'=============================================================================================================
'Capturar qual centro iria rodar batimento

resposta = InputBox("=============================" & vbCrLf & _
               "        Escolha uma opção:" & vbCrLf & _
               "=============================" & vbCrLf & _
               "1 - Batimento completo ( Nivel Inventario ) " & vbCrLf & _
               "2 - Batimento completo ( Nivel Diário )" & vbCrLf & _
               "3 - Batimento Status" & vbCrLf & _
               "4 - Batimento lote" & vbCrLf & _
               "5 - Sair" & vbCrLf & _
               "=============================")

'verificar se resposta foi vazia


If resposta = "" Then
    'selecionar o painel e ocultar as abas
    Ocultar
    ws_painel.Select
    MsgBox "Voce não selecionou nenhum tipo de opcão!!", vbInformation, "Batimento de Lote"
    Exit Sub
End If

' Verificar se a resposta está dentro das opções válidas (1 a 4)
    If Not IsNumeric(resposta) Or (resposta <> "1" And resposta <> "2" And resposta <> "3" And resposta <> "4") Then
        ' Se não for numérico ou estiver fora das opções, mostrar erro
        Ocultar
        ws_painel.Select
        MsgBox "Opção inválida! Selecione uma das opções: 1, 2, 3 ou 4.", vbCritical, "Batimento de Lote"
        Exit Sub
    End If

'=============================================================================================================
'CAPTURAR O CENTRO

resposta2 = InputBox("=============================" & vbCrLf & _
               "        Escolha uma opção:" & vbCrLf & _
               "=============================" & vbCrLf & _
               "1 - Manufatureiro - 7962" & vbCrLf & _
               "2 - Distribuidor - 7919" & vbCrLf & _
               "3 - Sair" & vbCrLf & _
               "=============================")

'VERIFICAR SE NAO ESTA EM BRANCO

If resposta2 = "" Then
    'selecionar o painel e ocultar as abas
    Ocultar
    ws_painel.Select
    MsgBox "Você não selecionou nenhum tipo de opcão!!", vbInformation, "Batimento de Lote"
    Exit Sub
End If


' Verificar se a resposta está dentro das opções válidas (1 a 4)
    If Not IsNumeric(resposta2) Or (resposta2 <> "1" And resposta2 <> "2" And resposta2 <> "3") Then
        ' Se não for numérico ou estiver fora das opções, mostrar erro
        Ocultar
        ws_painel.Select
        MsgBox "Opção inválida! Selecione uma das opções: 1, 2, ou 3.", vbCritical, "Batimento de Lote"
        Exit Sub
    End If

If resposta2 = 1 Then

cCentro = "Manufatureiro"
CENTRO = 7962

Else

cCentro = "Distribuidor"
CENTRO = 7919
End If
'=============================================================================================================

'VERIFICAR O TIPO DO SALDO

resposta3 = InputBox("=============================" & vbCrLf & _
               "        Escolha uma opção:" & vbCrLf & _
               "=============================" & vbCrLf & _
               "1 - Saldo do Robô" & vbCrLf & _
               "2 - Saldo Manual" & vbCrLf & _
               "3 - Sair" & vbCrLf & _
               "=============================")

If resposta3 = "" Then
    'selecionar o painel e ocultar as abas
    Ocultar
    ws_painel.Select
    
    MsgBox "Voce não selecionou nenhum tipo de opcão!!", vbInformation, "Batimento de Lote"
    Exit Sub
End If


' Verificar se a resposta está dentro das opções válidas (1 a 4)
    If Not IsNumeric(resposta3) Or (resposta3 <> "1" And resposta3 <> "2" And resposta3 <> "3") Then
        ' Se não for numérico ou estiver fora das opções, mostrar erro
        Ocultar
        ws_painel.Select
        MsgBox "Opção inválida! Selecione uma das opções: 1, 2, ou 3.", vbCritical, "Batimento de Lote"
        Exit Sub
    End If
'=============================================================================================================
'- pegar gpid
    Dim GetUserN
    Dim ObjNetwork
    Set ObjNetwork = CreateObject("WScript.Network")
    GetUserN = ObjNetwork.UserName
    UsuarioRede = GetUserN
    
If Not IsNumeric(UsuarioRede) Then ' em caso da maquina nao ser pepsico , nao puxa o gpid e pessoa necessita digitar o gpid
gpid = InputBox("Por favor, insira seu GPID (apenas números):", "Entrada de GPID")

' Verificar se o GPID é numérico
If IsNumeric(gpid) And Len(gpid) > 0 Then
UsuarioRede = gpid
Else
Ocultar
ws_painel.Select
MsgBox "Entrada inválida. O GPID deve ser composto apenas por números.", vbCritical, "Erro"
Exit Sub
End If
End If


eErros = 0
validationSaldo = False

'=============================================================================================================
'conexao SAP



'On Error GoTo SAP

 'CONEXÃO

    Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
    Set session = SAPCon.Children(0) 'Get the first session (window)


'=============================================================================================================
  ws_base.Select ' limpar nossa aba da base
  
  
  If ActiveSheet.FilterMode Then
    
    ActiveSheet.ShowAllData     ' Limpa todos os filtros da planilha
    
    End If
  
  wsBaseLN = ws_base.Cells(Rows.Count, "A").End(xlUp).Row
  
  If wsBaseLN > 5 Then
    ws_base.Range("A6:P" & wsBaseLN).ClearContents
  End If
  
  

'=============================================================================================================
   
' limpar a parte da validacao

ws_Arquivos.Select

lnValidacao = ws_Arquivos.Cells(Rows.Count, "bp").End(xlUp).Row

If lnValidacao > 1 Then

    ws_Arquivos.Range("Bp2:br" & lnValidacao).ClearContents
    
End If
'=============================================================================================================
 'limpar tabela de controle de execucao
ws_dados.Select

lndados = Application.WorksheetFunction.CountA(ws_dados.Range("A:A"))

If lndados > 1 Then
    ws_dados.Range("A2:I" & lndados).ClearContents
End If
'============================================================================================================= <======== batimento nivel inventario
    ' Verifica qual opção foi selecionada e qual o caminho a seguir
    
    Select Case resposta
        Case "1" 'Batimento completo nivel de inventario '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@caso 1
            
  '- > GERAR A BASE PARA STATUS -----------------------------------------------------------------------------------
  
  'puxar ZM
  
 tipZM = "STATUS"
 
 GERAR_ZM
 
  
sb_base_status_lote ' chamar mod para atualizar base de status

If validationSaldo Then
    Ocultar
    ws_painel.Select
    Exit Sub
 End If
 
formulas_da_base_STATUS ' INSERIMOS Os demais dados com formula
   
'=============================================================================================================
'atualizar dados
   

ws_dados.Select

lndados = Application.WorksheetFunction.CountA(ws_dados.Range("A:A"))

ws_dados.Range("A" & lndados + 1).Value = UsuarioRede

ws_dados.Range("b" & lndados + 1).Value = Format(DateValue(Now), "DD_MM_YYYY")

ws_dados.Range("c" & lndados + 1).Value = Format(Now, "hh_mm_ss")

ws_dados.Range("e" & lndados + 1).Value = "Batimento_Desbloqueio"

ws_dados.Range("G" & lndados + 1).Value = 0

ws_dados.Range("H" & lndados + 1).Value = 0
   
'=============================================================================================================
    ' limpar cabeçalho
    
    'filtrar
    
 ws_base.Select
 
 ws_base.Range("A2:p2").ClearContents
 
       
 '=============================================================================================================
    '  fitrar dados para batimento de desbloqueio e adicionar a seu dicionario
    
 
 ws_base.Range("P2").Value = "OK"
 ws_base.Range("N2").Value = "D"
 'ws_base.Range("E2").Value = "<>0"
  ws_base.Range("C2").Value = "<>GENERICO"
 
 Filtro_Avancado

  '=============================================================================================================
     session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1B"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "343"

    If resposta2 = 1 Then
        session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7962"
    Else
    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7919"
    End If

    session.findById("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
    
    
   
  '============================================================================================================
   wsBaseLN = ws_base.Cells(Rows.Count, "S").End(xlUp).Row
   
    V_aux = 0
   l_desbloqueio = 0
   eErros = 0
  If wsBaseLN > 5 Then 'verificar temos material a desbloquear
  
  ' Criar um objeto Dictionary para garantir que não haja duplicados
    Set dictDesbloqueio = CreateObject("Scripting.Dictionary")
    
   ' Percorrer a coluna S, a partir da linha 6
    For i = 6 To wsBaseLN
        valor = ws_base.Cells(i, "S").Value
        ' Adicionar o valor ao dicionário apenas se não for duplicado
        If Not dictDesbloqueio.exists(valor) And valor <> "" Then
            dictDesbloqueio.Add valor, Nothing
        End If
    Next i
    
    
   Set destino = ws_base.Range("B2") ' DESTINO onde vamos inserir os materias
    
    
    
    ' Inserir os valores únicos na coluna AX
    For Each valor In dictDesbloqueio.Keys
    
        destino.Value = valor 'inserimos o material
        Filtro_Avancado ' filtramos
        wsBaseLN = ws_base.Cells(Rows.Count, "S").End(xlUp).Row
    ' realizar o batimento de Status
     V_aux = wsBaseLN - 5
    sb_Batimento_Status_desbloqueio
    l_desbloqueio = l_desbloqueio + V_aux
        
    Next valor
    
    
End If


If eErros > 0 Then
l_desbloqueio = l_desbloqueio - eErros
End If


ws_dados.Select

ws_dados.Range("d" & lndados + 1).Value = Format(Now, "hh_mm_ss")

If wsBaseLN = 5 Then

    ws_dados.Range("f" & lndados + 1).Value = 0
   
Else

    ws_dados.Range("f" & lndados + 1).Value = l_desbloqueio
   
End If

ws_dados.Range("I" & lndados + 1).Value = ws_dados.Range("f" & lndados + 1).Value + ws_dados.Range("g" & lndados + 1).Value
    
    
 'ENTRAR ROTINA DE BLOQUEIO <-----------------------------------------
 'FILTRAR PARA BLOQUEIO
 'ADICIONAR AO DICIONARIO DE BLOQUEIO OS MATERIAS
    
  ws_base.Select
 
 ws_base.Range("A2:p2").ClearContents
 
ws_dados.Select

lndados = Application.WorksheetFunction.CountA(ws_dados.Range("A:A"))

ws_dados.Range("A" & lndados + 1).Value = UsuarioRede

ws_dados.Range("b" & lndados + 1).Value = Format(DateValue(Now), "DD_MM_YYYY")

ws_dados.Range("c" & lndados + 1).Value = Format(Now, "hh_mm_ss")

ws_dados.Range("e" & lndados + 1).Value = "Batimento_Bloqueio"

ws_dados.Range("G" & lndados + 1).Value = 0

ws_dados.Range("H" & lndados + 1).Value = 0

 '=============================================================================================================
    '  fitrar dados para batimento de BLOQUEIO e adicionar a seu dicionario
    
 ws_base.Select
 
 ws_base.Range("P2").Value = "OK"
 ws_base.Range("N2").Value = "B"
 'ws_base.Range("E2").Value = "<>0"
 ws_base.Range("C2").Value = "<>GENERICO"
 
 Filtro_Avancado


  '=============================================================================================================
   wsBaseLN = ws_base.Cells(Rows.Count, "S").End(xlUp).Row
   V_aux = 0
   l_bloqueio = 0
   eErros = 0
 If wsBaseLN > 5 Then ' verificar se temos dados para bloquear
 
  ' Criar um objeto Dictionary para garantir que não haja duplicados
    Set dictBloqueio = CreateObject("Scripting.Dictionary")
    
   ' Percorrer a coluna S, a partir da linha 6
    For i = 6 To wsBaseLN
        valor = ws_base.Cells(i, "S").Value
        ' Adicionar o valor ao dicionário apenas se não for duplicado
        If Not dictBloqueio.exists(valor) And valor <> "" Then
            dictBloqueio.Add valor, Nothing
        End If
    Next i
    
    
   Set destino = ws_base.Range("B2") ' DESTINO onde vamos inserir os materias
    
    
    
    ' Inserir os valores únicos na coluna AX
    For Each valor In dictBloqueio.Keys
    
        destino.Value = valor 'inserimos o material
        Filtro_Avancado ' filtramos
 
 'verificar uma possibilidade , se tiver 1 linha e nao tiver saldo no sap , nao tem oque bloquear
 

 ultimaLinha = ws_base.Cells(Rows.Count, "R").End(xlUp).Row
 
 If ultimaLinha = 6 Then
    If ws_base.Range("V6").Value <> 0 Then
        ' Executa a rotina de bloqueio se o valor de V6 não for zero
        V_aux = ultimaLinha - 5
        l_bloqueio = l_bloqueio + V_aux
        sb_Batimento_Status_BLOQUEIO
    End If
Else
    ' Caso haja mais de 6 linhas, verifica a coluna V
    possuiValorDiferenteDeZero = False
    For i = 6 To ultimaLinha
        If ws_base.Cells(i, "V").Value <> 0 Then
            possuiValorDiferenteDeZero = True
            Exit For ' Sai do loop se encontrar um valor diferente de zero
        End If
    Next i
    
    ' Se existir algum valor diferente de zero, executa a rotina de bloqueio
    If possuiValorDiferenteDeZero Then
        V_aux = ultimaLinha - 5
        l_bloqueio = l_bloqueio + V_aux
        sb_Batimento_Status_BLOQUEIO
    End If
End If
        
    Next valor

End If

If eErros > 0 Then
l_bloqueio = l_bloqueio - eErros
End If

ws_dados.Select

ws_dados.Range("d" & lndados + 1).Value = Format(Now, "hh_mm_ss")

If wsBaseLN = 5 Then

    ws_dados.Range("f" & lndados + 1).Value = 0
    
Else

    ws_dados.Range("f" & lndados + 1).Value = l_bloqueio
 
End If
  
  ws_dados.Range("I" & lndados + 1).Value = ws_dados.Range("f" & lndados + 1).Value + ws_dados.Range("g" & lndados + 1).Value
  
  '- > GERAR A BASE PARA LOTE
            
   '=============================================================================================================
  ws_base.Select ' limpar nossa aba da base
  
  
  If ActiveSheet.FilterMode Then
    
    ActiveSheet.ShowAllData     ' Limpa todos os filtros da planilha
    
    End If
  
  wsBaseLN = ws_base.Cells(Rows.Count, "A").End(xlUp).Row
  
  If wsBaseLN > 5 Then
    ws_base.Range("A6:P" & wsBaseLN).ClearContents
  End If
  
  

'=============================================================================================================
   
' limpar a parte da validacao

ws_Arquivos.Select

lnValidacao = ws_Arquivos.Cells(Rows.Count, "bp").End(xlUp).Row

If lnValidacao > 1 Then

    ws_Arquivos.Range("Bp2:br" & lnValidacao).ClearContents
    
End If
         
'=============================================================================================================
 'puxar ZM
tipZM = "LOTE"
 GERAR_ZM

sb_base_status_lote ' chamar mod para atualizar base de status

formulas_da_base_LOTE ' INSERIMOS Os demais dados com formula

 '=============================================================================================================
'atualizar dados
   

ws_dados.Select

lndados = Application.WorksheetFunction.CountA(ws_dados.Range("A:A"))

ws_dados.Range("A" & lndados + 1).Value = UsuarioRede

ws_dados.Range("b" & lndados + 1).Value = Format(DateValue(Now), "DD_MM_YYYY")

ws_dados.Range("c" & lndados + 1).Value = Format(Now, "hh_mm_ss")

ws_dados.Range("e" & lndados + 1).Value = "Batimento_Lote"
   
'=============================================================================================================

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1B"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "311"

    If resposta2 = 1 Then
    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7962"
    Else
    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7919"
    End If

    session.findById("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
    
  



'-=============================================================================================================
ws_base.Select
w…4229 tokens truncated…("AB5"), Order1:=xlAscending, Header:=xlYes
     
     If ws_base.AutoFilterMode Then ' retira filtro
        ws_base.AutoFilterMode = False
    End If
    ' realizar o batimento de lote

        V_aux = ultimaLinha - 5
        l_Lote = l_Lote + V_aux
         batimento_lote_nivel_diario
         
            
            
       End If
      
   End If
   
    Next valor
    
    If eErros > 0 Then
        l_Lote = l_Lote - eErros
    End If
    
ws_dados.Select

ws_dados.Range("d" & lndados + 1).Value = Format(Now, "hh_mm_ss")

ws_dados.Range("f" & lndados + 1).Value = l_Lote
ws_dados.Range("g" & lndados + 1).Value = l_Sobrante
ws_dados.Range("h" & lndados + 1).Value = l_faltante
  
ws_dados.Range("I" & lndados + 1).Value = ws_dados.Range("f" & lndados + 1).Value + ws_dados.Range("g" & lndados + 1).Value + ws_dados.Range("h" & lndados + 1).Value
  
    
End If
      
        Case "3" 'Batimento status @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@caso 3
        
        
  '=============================================================================================================
  '- > GERAR A BASE PARA STATUS -----------------------------------------------------------------------------------
  
  'puxar ZM
  
 tipZM = "STATUS"
 
 GERAR_ZM
 
  
sb_base_status_lote ' chamar mod para atualizar base de status

If validationSaldo Then
    Ocultar
    ws_painel.Select
    Exit Sub
 End If
 
formulas_da_base_STATUS ' INSERIMOS Os demais dados com formula
   
   
'=============================================================================================================
'atualizar dados
   

ws_dados.Select

lndados = Application.WorksheetFunction.CountA(ws_dados.Range("A:A"))

ws_dados.Range("A" & lndados + 1).Value = UsuarioRede

ws_dados.Range("b" & lndados + 1).Value = Format(DateValue(Now), "DD_MM_YYYY")

ws_dados.Range("c" & lndados + 1).Value = Format(Now, "hh_mm_ss")

ws_dados.Range("e" & lndados + 1).Value = "Batimento_Desbloqueio"

ws_dados.Range("G" & lndados + 1).Value = 0

ws_dados.Range("H" & lndados + 1).Value = 0

'=============================================================================================================
    ' limpar cabeçalho
    
    'filtrar
    
 ws_base.Select
 
 ws_base.Range("A2:p2").ClearContents
 
       
 '=============================================================================================================
    '  fitrar dados para batimento de desbloqueio e adicionar a seu dicionario
    
 
 ws_base.Range("P2").Value = "OK"
 ws_base.Range("N2").Value = "D"
 'ws_base.Range("E2").Value = "<>0"
  ws_base.Range("C2").Value = "<>GENERICO"
 
 Filtro_Avancado

  '=============================================================================================================
     session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1B"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "343"

    If resposta2 = 1 Then
        session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7962"
    Else
    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7919"
    End If

    session.findById("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
    
    
   
  '============================================================================================================
   wsBaseLN = ws_base.Cells(Rows.Count, "S").End(xlUp).Row
   
   V_aux = 0
   l_desbloqueio = 0
   eErros = 0
  If wsBaseLN > 5 Then 'verificar temos material a desbloquear
  
  ' Criar um objeto Dictionary para garantir que não haja duplicados
    Set dictDesbloqueio = CreateObject("Scripting.Dictionary")
    
   ' Percorrer a coluna S, a partir da linha 6
    For i = 6 To wsBaseLN
        valor = ws_base.Cells(i, "S").Value
        ' Adicionar o valor ao dicionário apenas se não for duplicado
        If Not dictDesbloqueio.exists(valor) And valor <> "" Then
            dictDesbloqueio.Add valor, Nothing
        End If
    Next i
    
    
   Set destino = ws_base.Range("B2") ' DESTINO onde vamos inserir os materias
    
    
    
    ' Inserir os valores únicos na coluna AX
    For Each valor In dictDesbloqueio.Keys
    
        destino.Value = valor 'inserimos o material
        Filtro_Avancado ' filtramos
        wsBaseLN = ws_base.Cells(Rows.Count, "S").End(xlUp).Row
    ' realizar o batimento de Status
    V_aux = wsBaseLN - 5
    sb_Batimento_Status_desbloqueio
        
    l_desbloqueio = l_desbloqueio + V_aux
    Next valor
    
    
End If

If eErros > 0 Then
l_desbloqueio = l_desbloqueio - eErros
End If

ws_dados.Select

ws_dados.Range("d" & lndados + 1).Value = Format(Now, "hh_mm_ss")

If wsBaseLN = 5 Then

    ws_dados.Range("f" & lndados + 1).Value = 0
   
Else

    ws_dados.Range("f" & lndados + 1).Value = l_desbloqueio
   
End If

    ws_dados.Range("I" & lndados + 1).Value = ws_dados.Range("f" & lndados + 1).Value + ws_dados.Range("g" & lndados + 1).Value

 'ENTRAR ROTINA DE BLOQUEIO <-----------------------------------------
 'FILTRAR PARA BLOQUEIO
 'ADICIONAR AO DICIONARIO DE BLOQUEIO OS MATERIAS
    
  ws_base.Select
 
 ws_base.Range("A2:p2").ClearContents
 
       
 '=============================================================================================================
 
 '=============================================================================================================
'atualizar dados
   

ws_dados.Select

lndados = Application.WorksheetFunction.CountA(ws_dados.Range("A:A"))

ws_dados.Range("A" & lndados + 1).Value = UsuarioRede

ws_dados.Range("b" & lndados + 1).Value = Format(DateValue(Now), "DD_MM_YYYY")

ws_dados.Range("c" & lndados + 1).Value = Format(Now, "hh_mm_ss")

ws_dados.Range("e" & lndados + 1).Value = "Batimento_Bloqueio"

ws_dados.Range("G" & lndados + 1).Value = 0

ws_dados.Range("H" & lndados + 1).Value = 0


    '  fitrar dados para batimento de BLOQUEIO e adicionar a seu dicionario
    
 ws_base.Select
 
 ws_base.Range("P2").Value = "OK"
 ws_base.Range("N2").Value = "B"
 'ws_base.Range("E2").Value = "<>0"
 ws_base.Range("C2").Value = "<>GENERICO"
 
 Filtro_Avancado


  '=============================================================================================================
   wsBaseLN = ws_base.Cells(Rows.Count, "S").End(xlUp).Row
   
  V_aux = 0
  l_bloqueio = 0
  eErros = 0
  
 If wsBaseLN > 5 Then ' verificar se temos dados para bloquear
 
  ' Criar um objeto Dictionary para garantir que não haja duplicados
    Set dictBloqueio = CreateObject("Scripting.Dictionary")
    
   ' Percorrer a coluna S, a partir da linha 6
    For i = 6 To wsBaseLN
        valor = ws_base.Cells(i, "S").Value
        ' Adicionar o valor ao dicionário apenas se não for duplicado
        If Not dictBloqueio.exists(valor) And valor <> "" Then
            dictBloqueio.Add valor, Nothing
        End If
    Next i
    
    
   Set destino = ws_base.Range("B2") ' DESTINO onde vamos inserir os materias
    
    
    
    ' Inserir os valores únicos na coluna AX
    For Each valor In dictBloqueio.Keys
    
        destino.Value = valor 'inserimos o material
        Filtro_Avancado ' filtramos
 
 'verificar uma possibilidade , se tiver 1 linha e nao tiver saldo no sap , nao tem oque bloquear
 

 ultimaLinha = ws_base.Cells(Rows.Count, "R").End(xlUp).Row
 
 If ultimaLinha = 6 Then
    If ws_base.Range("V6").Value <> 0 Then
        V_aux = ultimaLinha - 5
        l_bloqueio = l_bloqueio + V_aux
        ' Executa a rotina de bloqueio se o valor de V6 não for zero
        sb_Batimento_Status_BLOQUEIO
    End If
Else
    ' Caso haja mais de 6 linhas, verifica a coluna V
    possuiValorDiferenteDeZero = False
    For i = 6 To ultimaLinha
        If ws_base.Cells(i, "V").Value <> 0 Then
            possuiValorDiferenteDeZero = True
            Exit For ' Sai do loop se encontrar um valor diferente de zero
        End If
    Next i
    
    ' Se existir algum valor diferente de zero, executa a rotina de bloqueio
    If possuiValorDiferenteDeZero Then
         V_aux = ultimaLinha - 5
        l_bloqueio = l_bloqueio + V_aux
        sb_Batimento_Status_BLOQUEIO
    End If
End If
        
    Next valor

End If


If eErros > 0 Then
l_bloqueio = l_bloqueio - eErros
End If

ws_dados.Select

ws_dados.Range("d" & lndados + 1).Value = Format(Now, "hh_mm_ss")

If wsBaseLN = 5 Then

    ws_dados.Range("f" & lndados + 1).Value = 0
    
Else

    ws_dados.Range("f" & lndados + 1).Value = l_bloqueio
 
End If
  
  ws_dados.Range("I" & lndados + 1).Value = ws_dados.Range("f" & lndados + 1).Value + ws_dados.Range("g" & lndados + 1).Value
        
          
          
        Case "4" 'batimento lote @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@caso 4
           
'=============================================================================================================
 'puxar ZM
tipZM = "LOTE"
 GERAR_ZM

sb_base_status_lote ' chamar mod para atualizar base de status

If validationSaldo Then
    Ocultar
    ws_painel.Select
    Exit Sub
 End If
 
formulas_da_base_LOTE ' INSERIMOS Os demais dados com formula
   
 '=============================================================================================================
'atualizar dados
   

ws_dados.Select

lndados = Application.WorksheetFunction.CountA(ws_dados.Range("A:A"))

ws_dados.Range("A" & lndados + 1).Value = UsuarioRede

ws_dados.Range("b" & lndados + 1).Value = Format(DateValue(Now), "DD_MM_YYYY")

ws_dados.Range("c" & lndados + 1).Value = Format(Now, "hh_mm_ss")

ws_dados.Range("e" & lndados + 1).Value = "Batimento_Lote"

   
'=============================================================================================================

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1B"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "311"

    If resposta2 = 1 Then
    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7962"
    Else
    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7919"
    End If

'    session.findbyid("wnd[0]").sendVKey 0
'    session.findbyid("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
    
  



'-=============================================================================================================
ws_base.Select
ws_base.Range("A2:p2").ClearContents
 'ws_base.Range("P2").Value = "OK"
 ws_base.Range("k2").Value = "<>0"
 ws_base.Range("C2").Value = "<>GENERICO"
 
 Filtro_Avancado
   
'=============================================================================================================

 wsBaseLN = ws_base.Cells(Rows.Count, "S").End(xlUp).Row
   
  V_aux = 0
  l_Sobrante = 0
  l_faltante = 0
  l_Lote = 0
  eErros = 0
 If wsBaseLN > 5 Then 'verificar temos material a desbloquear
  
  ' Criar um objeto Dictionary para garantir que não haja duplicados
    Set dictLote = CreateObject("Scripting.Dictionary")
    
   ' Percorrer a coluna S, a partir da linha 6
    For i = 6 To wsBaseLN
        valor = ws_base.Cells(i, "S").Value
        ' Adicionar o valor ao dicionário apenas se não for duplicado
        If Not dictLote.exists(valor) And valor <> "" Then
            dictLote.Add valor, Nothing
        End If
    Next i
    
    
   Set destino = ws_base.Range("B2") ' DESTINO onde vamos inserir os materias
    
    
    
    ' Inserir os valores únicos na coluna AX
    For Each valor In dictLote.Keys
    
        destino.Value = valor 'inserimos o material
        Filtro_Avancado ' filtramos
        
         ultimaLinha = ws_base.Cells(Rows.Count, "R").End(xlUp).Row
         'validacao se todos lotes estao sobrando nao pode fazer batimento
         
         batSobrante = False
         batFaltante = False
 
            'rotina para verficar como esta o status do batimento /sobrando em todos , faltando ou dividido

                For i = 6 To ultimaLinha
                
                    valorBatimento = ws_base.Range("Y" & i).Value - ws_base.Range("v" & i)
                    
                    If valorBatimento < 0 Then
                    
                        batSobrante = True
                    Else
                        batFaltante = True
                  End If
                
                Next i
                
                    If batSobrante = True And batFaltante = False Then
                        
                        V_aux = ultimaLinha - 5
                        l_Sobrante = l_Sobrante + V_aux
                        ultimaLinha = 5
                        
                    ElseIf batSobrante = False And batFaltante = True Then
                        
                        V_aux = ultimaLinha - 5
                        l_faltante = l_faltante + V_aux
                        ultimaLinha = 5
                        
                    End If
            
    If ultimaLinha > 6 Then
    
    
     'Caso haja mais de 6 linhas, verifica a coluna V
    possuiValorDiferenteDeZero = False
    For i = 6 To ultimaLinha
        If ws_base.Cells(i, "V").Value <> 0 Then
            possuiValorDiferenteDeZero = True
            Exit For ' Sai do loop se encontrar um valor diferente de zero
        End If
    Next i
    
    If possuiValorDiferenteDeZero Then ' so realizar batimento se ao menos uma linha existir o saldo
    
     ' rotina para classificar do menor para maior
    ws_base.Range("R5:AG5").AutoFilter ' filtra

    ' classificar do maior para menor
     ws_base.Range("R5:AG" & ws_base.Cells(ws_base.Rows.Count, "R").End(xlUp).Row).Sort _
     Key1:=ws_base.Range("AB5"), Order1:=xlAscending, Header:=xlYes
     
     If ws_base.AutoFilterMode Then ' retira filtro
        ws_base.AutoFilterMode = False
    End If
    ' realizar o batimento de lote

        V_aux = ultimaLinha - 5
        l_Lote = l_Lote + V_aux
         batimento_lote_nivel_diario
         
            
            
       End If
      
   End If
   
    Next valor
    
    
 If eErros > 0 Then
        l_Lote = l_Lote - eErros
    End If
    
ws_dados.Select

ws_dados.Range("d" & lndados + 1).Value = Format(Now, "hh_mm_ss")

ws_dados.Range("f" & lndados + 1).Value = l_Lote
ws_dados.Range("g" & lndados + 1).Value = l_Sobrante
ws_dados.Range("h" & lndados + 1).Value = l_faltante
  
ws_dados.Range("I" & lndados + 1).Value = ws_dados.Range("f" & lndados + 1).Value + ws_dados.Range("g" & lndados + 1).Value + ws_dados.Range("h" & lndados + 1).Value
  
    
End If


        Case "5"  ' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@caso 5
            Ocultar
            ws_painel.Select
          Exit Sub      'sair
        Case Else
            Ocultar
            ws_painel.Select
            MsgBox "Opção inválida ,Macro cancelada", vbCritical + vbOKOnly, "Batimento de Lote" 'Qualquer outro tipo de opcao
        Exit Sub
    End Select



'LimparMemoriaExcel

MemoriaExcel

'envia email

e_mail_Batimento

' salvar dados no share point
share_dados_batimento

'ocultar abas
Ocultar
ws_painel.Select

'habilita tela

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.EnableEvents = True

MsgBox "Processo concluido!", vbInformation, "Batimento de Lotes -  NTV vs SAP"

Exit Sub

'SAP:

'MsgBox "Favor verificar sua conexão com SAP", vbCritical, "Erro de conexão com SAP"

End Sub



Sub GERAR_ZM()

Windows("Batimento Operacional.xlsm").Activate


ws_config.Select


caminho = ws_config.Range("A4").Value


On Error GoTo Sapv1
'CONEXÃO

Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
Set session = SAPCon.Children(0) 'Get the first session (window)

session.findById("wnd[0]").maximize


'session.findById("wnd[0]").resizeWorkingPane 147, 20, False
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZMXMM_MUNME"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = "300000000"
session.findById("wnd[0]/usr/ctxtMATNR-HIGH").Text = "399999999"
session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = CENTRO
session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = "IM01"
session.findById("wnd[0]/usr/ctxtCHARG-LOW").Text = "0000000000"
session.findById("wnd[0]/usr/ctxtCHARG-HIGH").Text = "9999999999"
session.findById("wnd[0]/usr/ctxtLGORT-LOW").SetFocus
session.findById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/mbar/menu[3]/menu[2]/menu[0]").Select
session.findById("wnd[1]/usr/btnAPP_FL_ALL").press
session.findById("wnd[1]/usr/btnB_SEARCH").press
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Material"
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 8
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/btnAPP_WL_SING").press
session.findById("wnd[1]/usr/btnB_SEARCH").press
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Texto breve material"
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 20
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/btnAPP_WL_SING").press
session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").Columns.elementAt(0).Width = 21
session.findById("wnd[1]/usr/btnB_SEARCH").press
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Lote"
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 4
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/btnAPP_WL_SING").press
session.findById("wnd[1]/usr/btnB_SEARCH").press
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Stock Unrest. (UMA1)"
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 18
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/btnAPP_WL_SING").press
session.findById("wnd[1]/usr/btnB_SEARCH").press
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Blocked Stock (UMA1)"
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 9
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/btnAPP_WL_SING").press
session.findById("wnd[1]/usr/btnB_SEARCH").press
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Restricted Stock (UMA1)"
session.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 21
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/btnAPP_WL_SING").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/ctxtDY_PATH").Text = caminho
session.findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "SALDO_SAP-" & tipZM & ".TXT"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[2]").sendVKey 11
session.findById("wnd[1]/tbar[0]/btn[11]").press

Exit Sub

Sapv1:

MsgBox "Sap esta fechado , favor manter o mesmo aberto!!!!", vbCritical + vbSystemModal, "BATIMENTO"


End Sub



