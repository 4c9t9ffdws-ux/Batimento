Attribute VB_Name = "MOD_BATIMENTO_LOTE_NIVEL_INVENT"
Sub batimento_lote_nivel_INV()

Dim lastow  As Integer, index   As Integer, valorBatimento As Integer, valorAux, indexAux, ln As Integer, indexSAP As Integer, cotroleItem As Integer, varQTD As Integer, ccLinha As Integer
Dim batSobrante As Boolean, batFaltante As Boolean, varNewItem As Boolean, errorT As Boolean


Dim dadosBatimento()    As Variant


' Obter a última linha de dados
lastRow = ws_base.Cells(Rows.Count, "V").End(xlUp).Row
ReDim dadosBatimento(1 To lastRow * 2, 1 To 7)

varNewItem = False

'CONEXAO SAP
    Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
    Set session = SAPCon.Children(0) 'Get the first session (window)
    
     session.findById("wnd[0]").sendVKey 0 ' ENTRAR NA TRANSACAO
     
     On Error Resume Next

session.findById("wnd[0]/usr/ctxtMSEGK-UMLGO").Text = "IM01"

On Error GoTo 0
 
batSobrante = False
batFaltante = False
ccLinha = 1
lnNovaM = 1

' Coletar dados e calcular diferenças
ln = 6
For index = 1 To lastRow - 5
    dadosBatimento(index, 1) = ws_base.Cells(ln, "S").Value ' MATERIAL
    dadosBatimento(index, 2) = ws_base.Cells(ln, "U").Value ' LOTE
    dadosBatimento(index, 3) = ws_base.Cells(ln, "V").Value ' SAP Livre
    dadosBatimento(index, 4) = ws_base.Cells(ln, "Y").Value ' DLX Livre
    ln = ln + 1
    controleItem = controleItem + 1
Next index

'rotina para verficar como esta o status do batimento /sobrando em todos , faltando ou dividido

For i = 1 To controleItem

    valorBatimento = dadosBatimento(i, 4) - dadosBatimento(i, 3)
    
    If valorBatimento < 0 Then
    
    batSobrante = True
    Else
    batFaltante = True
  End If

Next i


For index = 1 To controleItem

valorBatimento = dadosBatimento(index, 4) - dadosBatimento(index, 3)

' se valor batimento negativo esta sobrando
'se valor batimento posito esta faltando
    If valorBatimento > 0 Then 'faltando '**************************************IF principal
    
        'buscar sobrante
        ' indexAux = index + 1 ' inciar a busca na proxima linha
         
'********************************************************************************************************
              For indexAux = index + 1 To controleItem
                
                If dadosBatimento(indexAux, 3) = 0 Then ' significa que o lote esta sem saldo no sap , precisamos buscar em outra linha
                
                    For lnAUX = indexAux To controleItem
                    
                       If lnAUX = controitem Then
                       Else
                        indexAux = indexAux + 1
                       End If
                        If dadosBatimento(indexAux, 3) <> 0 Then
                            lnAUX = controleItem
                        End If
                        
                    Next lnAUX
                
                End If
                
                valorAux = dadosBatimento(indexAux, 4) - dadosBatimento(indexAux, 3)
                
                'se valor negativo , tem sobrando
                    If valorAux < 0 Then
                        'verificar se saldo atente e precisamos transformar valor em positivo
                       valorAux = valorAux * -1
            
                     
                       
                        If valorAux = valorBatimento Then  ' quantidade e exata
                        
                         dadosBatimento(index, 5) = valorAux
                         dadosBatimento(index, 6) = dadosBatimento(indexAux, 2)
                         dadosBatimento(indexAux, 3) = dadosBatimento(indexAux, 3) - valorAux
                         dadosBatimento(index, 7) = dadosBatimento(index, 2)
                         indexAux = controleItem
                          
                        ElseIf valorAux > valorBatimento Then ' quantidade e maior
                        
                        dadosBatimento(index, 3) = dadosBatimento(index, 3) + valorBatimento
                        dadosBatimento(indexAux, 3) = dadosBatimento(indexAux, 3) - valorBatimento
                        dadosBatimento(index, 5) = valorBatimento
                        dadosBatimento(index, 6) = dadosBatimento(indexAux, 2) ' lote retirar
                        dadosBatimento(index, 7) = dadosBatimento(index, 2) ' lote receber
                        
                        indexAux = controleItem
                        
                        Else ' quantidade e menor
                        
                        'validar contagem e buscar mais saldo se possivel ---------------------<
                        dadosBatimento(indexAux, 3) = dadosBatimento(indexAux, 3) - valorAux
                        dadosBatimento(index, 3) = dadosBatimento(index, 3) + valorAux
                        dadosBatimento(index, 5) = valorAux
                        dadosBatimento(index, 6) = dadosBatimento(indexAux, 2)
                        dadosBatimento(index, 7) = dadosBatimento(index, 2)
                       ' novo loop para buscar sobrantes
                       'criar nova linha na matriz
                        valorBatimento = dadosBatimento(index, 4) - dadosBatimento(index, 3)
                        
                        If lnNovaM = 1 Then
                        
                        lnNovaM = controleItem + 1 ' olhar aqui pq no segundo loop ela vai incrementar errado , entao precisa definir  ela fora daqui
                        
                       End If
                        
                        If valorBatimento <> 0 Then ' buscar mais saldos sobrantes
                        
                            For j = indexAux + 1 To controleItem
                                
                                valorAux = dadosBatimento(j, 4) - dadosBatimento(j, 3)
                                
                                 If valorAux < 0 Then
                                    valorAux = valorAux * -1
                                 End If
                                 
                                  If valorAux < valorBatimento And j <= controleItem And valorAux <> 0 Then
                                  
                                    dadosBatimento(lnNovaM, 1) = dadosBatimento(index, 1)
                                    dadosBatimento(lnNovaM, 2) = dadosBatimento(index, 2)
                                    dadosBatimento(j, 3) = dadosBatimento(j, 3) - valorAux
                                    dadosBatimento(lnNovaM, 3) = dadosBatimento(index, 3) + valorAux
                                    dadosBatimento(lnNovaM, 4) = dadosBatimento(index, 4)
                                    dadosBatimento(lnNovaM, 5) = valorAux
                                    dadosBatimento(lnNovaM, 6) = dadosBatimento(j, 2)
                                    dadosBatimento(lnNovaM, 7) = dadosBatimento(index, 2)
                                    dadosBatimento(index, 3) = dadosBatimento(index, 3) + valorAux
                                                
                                    valorBatimento = dadosBatimento(index, 4) - dadosBatimento(index, 3)
                                  Else
                                  
                                    dadosBatimento(lnNovaM, 1) = dadosBatimento(index, 1)
                                    dadosBatimento(lnNovaM, 2) = dadosBatimento(index, 2)
                                    
                                    If valorBatimento <= dadosBatimento(j, 3) Then
                                    
                                        dadosBatimento(j, 3) = dadosBatimento(j, 3) - valorBatimento
                                        dadosBatimento(lnNovaM, 3) = dadosBatimento(index, 3) + valorBatimento
                                        dadosBatimento(index, 3) = dadosBatimento(index, 3) + valorBatimento
                                    
                                    Else
                                        
                                        dadosBatimento(lnNovaM, 3) = dadosBatimento(index, 3) + valorBatimento
                                        valorBatimento = dadosBatimento(j, 3)
                                        dadosBatimento(j, 3) = dadosBatimento(j, 3) - dadosBatimento(j, 3)
                                        dadosBatimento(index, 3) = dadosBatimento(index, 3) + valorBatimento
                                    End If
                                    
                                    
                                    dadosBatimento(lnNovaM, 4) = dadosBatimento(index, 4)
                                    dadosBatimento(lnNovaM, 5) = valorBatimento
                                    dadosBatimento(lnNovaM, 6) = dadosBatimento(j, 2)
                                    dadosBatimento(lnNovaM, 7) = dadosBatimento(index, 2)
                                    
                                                
                                    valorBatimento = dadosBatimento(index, 4) - dadosBatimento(index, 3)
                                  
                                    If valorBatimento = 0 Then
                                        j = controleItem
                                    End If
                                  
                                  End If
                            
                                lnNovaM = lnNovaM + 1
                            Next j
                        
                        End If
                        
                        indexAux = controleItem
                        
                        'breaks = 22
                        
                        
                        
                        End If
                        
                    
                    Else ' nao ta sobrando  , precisamos matar maximo de linhas
                    
                    If dadosBatimento(indexAux, 3) <> 0 And dadosBatimento(indexAux, 3) >= valorBatimento Then
                       
                        dadosBatimento(indexAux, 3) = dadosBatimento(indexAux, 3) - valorBatimento
                        dadosBatimento(index, 3) = dadosBatimento(index, 3) + valorBatimento
                        dadosBatimento(index, 5) = valorBatimento
                        dadosBatimento(index, 6) = dadosBatimento(indexAux, 2)
                        dadosBatimento(index, 7) = dadosBatimento(index, 2)
                    
                    indexAux = controleItem
                    
                    Else ' fazer quando zerado ir outras linhas
                    
                    If dadosBatimento(indexAux, 3) > 0 Then
                    
                    dadosBatimento(index, 3) = dadosBatimento(index, 3) + dadosBatimento(indexAux, 3)
                    valorBatimento = dadosBatimento(indexAux, 3)
                    dadosBatimento(indexAux, 3) = dadosBatimento(indexAux, 3) - dadosBatimento(indexAux, 3)
                    dadosBatimento(index, 5) = valorBatimento
                    dadosBatimento(index, 6) = dadosBatimento(indexAux, 2)
                    dadosBatimento(index, 7) = dadosBatimento(index, 2)
                   End If
                     valorBatimento = dadosBatimento(index, 4) - dadosBatimento(index, 3)
                        
                       If lnNovaM = 1 Then
                        
                        lnNovaM = controleItem + 1 ' olhar aqui pq no segundo loop ela vai incrementar errado , entao precisa definir  ela fora daqui
                        
                       End If
                       
                     If valorBatimento <> 0 Then ' buscar mais saldos sobrantes
                     
                     
                        For j = indexAux + 1 To controleItem
                        
                        If valorBatimento <= dadosBatimento(j, 3) Then
                        
                                    dadosBatimento(lnNovaM, 1) = dadosBatimento(index, 1)
                                    dadosBatimento(lnNovaM, 2) = dadosBatimento(index, 2)
                                    dadosBatimento(j, 3) = dadosBatimento(j, 3) - valorBatimento
                                    dadosBatimento(lnNovaM, 3) = dadosBatimento(index, 3) + valorBatimento
                                    dadosBatimento(lnNovaM, 4) = dadosBatimento(index, 4)
                                    dadosBatimento(lnNovaM, 5) = valorBatimento
                                    dadosBatimento(lnNovaM, 6) = dadosBatimento(j, 2)
                                    dadosBatimento(lnNovaM, 7) = dadosBatimento(index, 2)
                                    dadosBatimento(index, 3) = dadosBatimento(index, 3) + valorBatimento
                                    j = controleItem
                                    indexAux = controleItem
                                    'lnNovaM = lnNovaM - 1
                        
                        Else
                                    
                                 If dadosBatimento(j, 3) > 0 Then
                                 
                                    dadosBatimento(lnNovaM, 1) = dadosBatimento(index, 1)
                                    dadosBatimento(lnNovaM, 2) = dadosBatimento(index, 2)
                                    valorBatimento = dadosBatimento(j, 3)
                                    dadosBatimento(lnNovaM, 3) = dadosBatimento(index, 3) + dadosBatimento(j, 3)
                                    dadosBatimento(j, 3) = dadosBatimento(j, 3) - dadosBatimento(j, 3)
                                    dadosBatimento(lnNovaM, 4) = dadosBatimento(index, 4)
                                    dadosBatimento(lnNovaM, 5) = valorBatimento
                                    dadosBatimento(lnNovaM, 6) = dadosBatimento(j, 2)
                                    dadosBatimento(lnNovaM, 7) = dadosBatimento(index, 2)
                                    dadosBatimento(index, 3) = dadosBatimento(index, 3) + valorBatimento
                                    valorBatimento = dadosBatimento(index, 4) - dadosBatimento(index, 3)
                                  End If
                                    If valorBatimento = 0 Then
                                        j = controleItem
                                        indexAux = controleItem
                                    End If
                        
                        End If
                            'valorAux = dadosBatimento(j, 4) - dadosBatimento(j, 3)
                        If dadosBatimento(j, 3) > 0 Then
                        
                          lnNovaM = lnNovaM + 1
                        End If
                        
                        Next j
                     
                     End If
                    
                    End If
                    
                    
              
                    
                End If
                
                
                
               
                
            Next indexAux
            
'***************************************************************************************************************
    ElseIf valorBatimento < 0 Then ' sobrando ****************************************2 condição
    
      If batSobrante = True And batFaltante = True And index = controleItem Then
    
'        dadosBatimento(index, 5) = 0
'        dadosBatimento(index, 6) = 0
        
       Else ' distribuir saldo sobrante
             
       valorBatimento = valorBatimento * -1
            
       If index < controleItem Then ' so fazer se nao for ultimo item
       
       dadosBatimento(index, 3) = dadosBatimento(index, 3) - valorBatimento
       dadosBatimento(index + 1, 3) = dadosBatimento(index + 1, 3) + valorBatimento
       dadosBatimento(index, 5) = valorBatimento
       dadosBatimento(index, 6) = dadosBatimento(index, 2)
       dadosBatimento(index, 7) = dadosBatimento(index + 1, 2)
       End If
            
      breask = 0
      End If
    
    
    Else ' valor zerado  , nao precisa mexer *********************************************3 condição
    
    
    
  If batSobrante = True And batFaltante = True And valorBatimento <> 0 Then
  
    dadosBatimento(index, 5) = 0
    dadosBatimento(index, 6) = 0
  
  End If
    
End If
    
'lnNovaM = lnNovaM - 1
Next index


If controleItem < lnNovaM Then

    controleItem = lnNovaM
End If

'==========teste



'linhaSaida = 6  ' Ajuste para imprimir corretamente
'
'    For i = 1 To controleItem
'  ' If dadosBatimento(i, 5) <> "" Then
'    ' Imprimir resultados ajustados
'    ws_base.Cells(linhaSaida, 36).Value = dadosBatimento(i, 1) ' Material
'    ws_base.Cells(linhaSaida, 37).Value = dadosBatimento(i, 2) ' Lote
'    ws_base.Cells(linhaSaida, 38).Value = dadosBatimento(i, 3) ' SAP Livre
'    ws_base.Cells(linhaSaida, 39).Value = dadosBatimento(i, 4) ' DLX Livre
'    ws_base.Cells(linhaSaida, 40).Value = dadosBatimento(i, 5) ' Qtd Batimento
'    ws_base.Cells(linhaSaida, 41).Value = dadosBatimento(i, 6) ' Lote
'    ws_base.Cells(linhaSaida, 42).Value = dadosBatimento(i, 7) ' Lote
'    linhaSaida = linhaSaida + 1
' ' End If
'
'    Next i
'
'  ws_base.Range("AJ6").Select
'==========================================


errorT = False

'---------


 lnsap = session.findById("wnd[0]/usr/sub:SAPMM07M:0421").loopRowCount ' CONTAR Quantas linhas temos na tela sap
 
 contador = 0
 linhaSAP = 0
   'verificar quantos itens temos
   
   For ln = 1 To UBound(dadosBatimento, 1)
   
    If dadosBatimento(ln, 1) <> "" Then
    
        contador = contador + 1
        
    End If
   
   Next ln


If lnsap >= contador Then ' tenho qtd suficiente para 1 pagina

'loop para inserir itens
    
          For i = 1 To contador ' item 1 ate a quantidade de itens que esta na variavel contador
          
          
           'somente inserir valores contidos
           
               If dadosBatimento(i, 5) <> 0 Then
              
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[" & linhaSAP & ",7]").Text = dadosBatimento(i, 1)
                       
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[" & linhaSAP & ",26]").Text = dadosBatimento(i, 5)
                
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[" & linhaSAP & ",44]").Text = "CX"
                
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[" & linhaSAP & ",53]").Text = "000" & dadosBatimento(i, 6)
        
                linhaSAP = linhaSAP + 1
               End If
               
          Next i


i = 1
       session.findById("wnd[0]").sendVKey 0
             
      texto = session.findById("wnd[0]/sbar/pane[0]").Text
      
            Do Until texto = ""
            
            lotep = session.findById("wnd[0]/usr/ctxtMSEG-CHARG").Text
           If dadosBatimento(i, 7) <> "" Then
          
            session.findById("wnd[0]/usr/ctxtMSEG-UMCHA").Text = "000" & dadosBatimento(ccLinha, 7)
            
             session.findById("wnd[0]").sendVKey 0
             texto = session.findById("wnd[0]/sbar/pane[0]").Text
                 

                 Do Until varNewItem = True
                 
                    If InStr(texto, "O lote recebedora é definido automaticamente") > 0 Then ' lote transferido , incrementar linha mamtriz
                        'session.findbyid("wnd[0]").sendVKey 0
                        texto = ""
                        ccLinha = ccLinha + 1
                        varNewItem = True
                     Else
                        
                            If InStr(texto, "LO Livre utilização não atingido") > 0 Or InStr(texto, "não existe") > 0 Then
                            errorT = True
                            session.findById("wnd[0]").maximize
                            session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1B"
                            session.findById("wnd[0]").sendVKey 0
                            session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "311"
                            
                                If resposta2 = 1 Then
                                    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7962"
                                Else
                                session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7919"
                                End If
                            
                            'session.findbyid("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
                            texto = ""
                            varNewItem = True
                        End If
                    
                        If errorT = False Then
                        
                        session.findById("wnd[0]").sendVKey 0
                        texto = session.findById("wnd[0]/sbar/pane[0]").Text
                        
                        End If
                        
                            If texto = "" Then
                                
                                varNewItem = True
                             End If
                             
'
                    End If
                    
                   
                 
                 Loop

                 varNewItem = False
                 
                 
                 texto = session.findById("wnd[0]/sbar/pane[0]").Text
                 
                    If InStr(texto, "0 peças") > 0 Then
                        session.findById("wnd[0]").sendVKey 0
                        texto = ""
                        
                    End If
                    
                        If InStr(texto, "LO Livre utilização não atingido") > 0 Or InStr(texto, "não existe") > 0 Then
                        errorT = True
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
                        texto = ""
                        
                        End If
                   End If
                    i = i + 1
                        If dadosBatimento(i, 7) = "" Then
                         ccLinha = ccLinha + 1
                        End If
                Loop

Else ' vai precisar de mais de uma pagina

linhaSAP = 0
controle = 0
For i = 1 To contador

          
      
           'somente inserir valores contidos
           
               If dadosBatimento(i, 5) <> 0 Then
              
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[" & linhaSAP & ",7]").Text = dadosBatimento(i, 1)
                       
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[" & linhaSAP & ",26]").Text = dadosBatimento(i, 5)
                
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[" & linhaSAP & ",44]").Text = "CX"
                
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[" & linhaSAP & ",53]").Text = "000" & dadosBatimento(i, 6)
        
                linhaSAP = linhaSAP + 1
                controle = controle + 1
                
                 If controle = lnsap Or i = contador Then ' verificar transferencia e criar new page
                 
                    session.findById("wnd[0]").sendVKey 0
             
                    texto = session.findById("wnd[0]/sbar/pane[0]").Text
                 
                    Do Until texto = ""
            
                     lotep = session.findById("wnd[0]/usr/ctxtMSEG-CHARG").Text
                         If dadosBatimento(i, 7) <> "" Then
                         
                            session.findById("wnd[0]/usr/ctxtMSEG-UMCHA").Text = "000" & dadosBatimento(ccLinha, 7)
            
                            session.findById("wnd[0]").sendVKey 0
                            texto = session.findById("wnd[0]/sbar/pane[0]").Text
                            
                                Do Until varNewItem = True
                 
                                    If InStr(texto, "O lote recebedora é definido automaticamente") > 0 Then ' lote transferido , incrementar linha mamtriz
                                        'session.findbyid("wnd[0]").sendVKey 0
                                        'texto = ""
                                        ccLinha = ccLinha + 1
                                        varNewItem = True
                                     Else
                                        
                                        If InStr(texto, "LO Livre utilização não atingido") > 0 Or InStr(texto, "não existe") > 0 Then
                                            errorT = True
                                            session.findById("wnd[0]").maximize
                                            session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1B"
                                            session.findById("wnd[0]").sendVKey 0
                                            session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "311"
                                            
                                                If resposta2 = 1 Then
                                                    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7962"
                                                Else
                                                session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7919"
                                                End If
                                            
                                            'session.findbyid("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
                                            texto = ""
                                            varNewItem = True
                                            i = contador
                                            eErros = eErros + V_aux
                                    End If
                                    
                                       If errorT = False Then
                                       
                                        session.findById("wnd[0]").sendVKey 0
                                        texto = session.findById("wnd[0]/sbar/pane[0]").Text
                                        
                                        End If
                                        
                                            If texto = "" Then
                                                
                                                varNewItem = True
                                             End If
                                    End If
                                 
                                    
                    
                    
                                 Loop
                
                                 varNewItem = False
                               
                         End If
                         
                         
                   Loop
                 
                     If errorT = False Then
                 
                   linhaSAP = 0
                   controle = 0
                   session.findById("wnd[0]/tbar[1]/btn[19]").press
                 End If
                 
                   
                 End If
                   
                    
               End If
               
        


Next i




End If

If errorT = False Then

session.findById("wnd[0]/tbar[0]/btn[11]").press

Else

errorT = False

End If
ws_base.Range("Aj6").CurrentRegion.ClearContents

End Sub

