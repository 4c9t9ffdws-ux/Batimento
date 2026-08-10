Attribute VB_Name = "MOD_BATIMENTO_LOTE_NIVEL_DIARIO"
Sub batimento_lote_nivel_diario()

    Dim lastRow As Long, ln As Long, i As Long, j As Long
    Dim controleItem As Long, valorNecessario As Double, valorDisponivel As Double
    Dim valorTransferido As Double, lnNovaM As Long
    Dim dadosBatimento() As Variant
    Dim batSobrante As Boolean, batFaltante As Boolean, errorT As Boolean


    Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
    Set session = SAPCon.Children(0) 'Get the first session (window)
    

     session.findById("wnd[0]").sendVKey 0 ' ENTRAR NA TRANSACAO
     
     session.findById("wnd[0]/usr/ctxtMSEGK-UMLGO").Text = "IM01"
     
varNewItem = False
errorT = False
    ' Obter a última linha de dados
    lastRow = ws_base.Cells(ws_base.Rows.Count, "V").End(xlUp).Row
    ReDim dadosBatimento(1 To lastRow * 2 - 5, 1 To 7) ' Ajuste baseado no intervalo esperado
    
    ' Alimentar a matriz inicial
    ln = 6
    controleItem = 0
    For i = 1 To lastRow - 5
        dadosBatimento(i, 1) = ws_base.Cells(ln, "S").Value ' MATERIAL
        dadosBatimento(i, 2) = ws_base.Cells(ln, "U").Value ' LOTE
        dadosBatimento(i, 3) = ws_base.Cells(ln, "V").Value ' SAP Livre
        dadosBatimento(i, 4) = ws_base.Cells(ln, "Y").Value ' DLX Livre
        ln = ln + 1
        controleItem = controleItem + 1
    Next i
    
    ' Verificar o status geral do batimento
    batSobrante = False
    batFaltante = False
    For i = 1 To controleItem
        If dadosBatimento(i, 4) - dadosBatimento(i, 3) < 0 Then
            batSobrante = True
        Else
            batFaltante = True
        End If
    Next i
    
lnNovaM = controleItem

For i = 1 To controleItem
    ' Calcular o valor necessário para cada linha (DLX Livre - SAP Livre)
    valorNecessario = dadosBatimento(i, 4) - dadosBatimento(i, 3) ' DLX Livre - SAP Livre

     'localizar linha sobrante de baixo para cima
      '
      saldoSobrante = 0
     If valorNecessario > 0 Then
      
        For j = controleItem To 1 Step -1 ' Percorrer de baixo para cima
        ' Calcula o saldo disponível na linha (SAP Livre - DLX Livre)
        valorDisponivel = dadosBatimento(j, 3) - dadosBatimento(j, 4) ' SAP Livre - DLX Livre
        
        ' Se houver saldo sobrante (valor disponível maior que zero)
        If valorDisponivel > 0 Then
            ' Armazena o saldo sobrante encontrado
            saldoSobrante = valorDisponivel
            
            ' Sai do loop assim que encontrar o saldo sobrante
            Exit For
        End If
        Next j
        
     If saldoSobrante > 0 Then
     
        If valorDisponivel >= valorNecessario Then
        
           dadosBatimento(i, 5) = valorNecessario ' Quantidade transferida
           dadosBatimento(i, 6) = dadosBatimento(j, 2) ' Lote que vai retirar
           dadosBatimento(i, 7) = dadosBatimento(i, 2) ' Lote de origem
           dadosBatimento(j, 3) = dadosBatimento(j, 3) - valorNecessario ' Atualiza saldo SAP Livre
           dadosBatimento(i, 3) = dadosBatimento(i, 3) + valorNecessario ' Atualiza saldo SAP Livre
           
           valorNecessario = 0 ' O saldo necessário foi preenchido, então sai do
        Else ' nao tem saldo suficiente
        
           dadosBatimento(i, 5) = valorDisponivel ' Quantidade transferida
           dadosBatimento(i, 6) = dadosBatimento(j, 2) ' Lote que vai retirar
           dadosBatimento(i, 7) = dadosBatimento(i, 2) ' Lote de origem
           dadosBatimento(j, 3) = dadosBatimento(j, 3) - valorDisponivel ' Atualiza saldo SAP Livre
           dadosBatimento(i, 3) = dadosBatimento(i, 3) + valorDisponivel ' Atualiza saldo SAP Livre
           valorNecessario = dadosBatimento(i, 4) - dadosBatimento(i, 3)
           
           
           ' loop criando linha na matriz
           
              'For mov = controleItem To i + 1 Step -1
              For mov = 1 To controleItem
                valorDisponivel = dadosBatimento(mov, 4) - dadosBatimento(mov, 3) ' DLX Livre - SAP Livre
                
                    If valorDisponivel < 0 Then
                        lnNovaM = lnNovaM + 1
                        valorDisponivel = valorDisponivel * -1
                        
                         If valorDisponivel >= valorNecessario Then
                            dadosBatimento(lnNovaM, 1) = dadosBatimento(i, 1) ' material
                            dadosBatimento(lnNovaM, 2) = dadosBatimento(i, 2)  'lote
                            dadosBatimento(lnNovaM, 5) = valorNecessario ' Quantidade transferida
                            dadosBatimento(lnNovaM, 6) = dadosBatimento(mov, 2) ' Lote que vai retirar
                            dadosBatimento(lnNovaM, 7) = dadosBatimento(i, 2) ' Lote de origem
                            dadosBatimento(mov, 3) = dadosBatimento(mov, 3) - valorNecessario ' Atualiza saldo SAP Livre
                            dadosBatimento(i, 3) = dadosBatimento(i, 3) + valorNecessario ' Atualiza saldo SAP Livre
                            valorNecessario = dadosBatimento(i, 4) - dadosBatimento(i, 3)
                         
                            Exit For
                         Else
                          
                            dadosBatimento(lnNovaM, 1) = dadosBatimento(i, 1) ' material
                            dadosBatimento(lnNovaM, 2) = dadosBatimento(i, 2)  'lote
                            dadosBatimento(lnNovaM, 5) = valorDisponivel ' Quantidade transferida
                            dadosBatimento(lnNovaM, 6) = dadosBatimento(mov, 2) ' Lote que vai retirar
                            dadosBatimento(lnNovaM, 7) = dadosBatimento(i, 2) ' Lote de origem
                            dadosBatimento(mov, 3) = dadosBatimento(mov, 3) - valorDisponivel ' Atualiza saldo SAP Livre
                            dadosBatimento(i, 3) = dadosBatimento(i, 3) + valorDisponivel ' Atualiza saldo SAP Livre
                            valorNecessario = dadosBatimento(i, 4) - dadosBatimento(i, 3)
                        
                         End If
                    End If
              Next mov
        
        End If
     End If
    End If
Next i




If controleItem < lnNovaM Then

    controleItem = lnNovaM
End If

'==========teste


'linhaSaida = 6  ' Ajuste para imprimir corretamente
'
'    For i = 1 To controleItem
'
'   'If dadosBatimento(i, 5) <> "" Then
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

'===============================================================================



'---------


 lnsap = session.findById("wnd[0]/usr/sub:SAPMM07M:0421").loopRowCount ' CONTAR Quantas linhas temos na tela sap
 
 contador = 0
 linhaSAP = 0
 ccLinha = 1
   'verificar quantos itens temos
   
   For ln = 1 To UBound(dadosBatimento, 1)
   
    If dadosBatimento(ln, 1) <> "" Then
    
        contador = contador + 1
        
    End If
   
   Next ln

'rotina para controle da linha
contador2 = 0
 For ln = 1 To UBound(dadosBatimento, 1)
   
    If dadosBatimento(ln, 5) <> "" Then
    
        contador2 = contador2 + 1
        
    End If
   
   Next ln

If lnsap >= contador2 Then ' tenho qtd suficiente para 1 pagina

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
          
            session.findById("wnd[0]/usr/ctxtMSEG-UMCHA").Text = "000" & dadosBatimento(i, 7)
            
             session.findById("wnd[0]").sendVKey 0
             texto = session.findById("wnd[0]/sbar/pane[0]").Text
                 

                 Do Until varNewItem = True
                 
                    If InStr(texto, "O lote recebedora é definido automaticamente") > 0 Then ' lote transferido , incrementar linha mamtriz
                        'session.findbyid("wnd[0]").sendVKey 0
                        texto = ""
                        'ccLinha = ccLinha + 1
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
                                eErros = eErros + V_aux
                            End If
                    
                     If errorT = False Then
                    
                        session.findById("wnd[0]").sendVKey 0
                        texto = session.findById("wnd[0]/sbar/pane[0]").Text
                        
                     End If
                            If texto = "" Then
                                
                                varNewItem = True
                             End If
                             
'                                 If InStr(texto, "  ") > 0 Then
'
'                                 End If
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
                        
                        'session.findbyid("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
                        texto = ""
                        varNewItem = True
                        eErros = eErros + V_aux
                         
                    End If
                   End If
                   
'                    If dadosBatimento(i, 7) = "" Then
'                    ccLinha = ccLinha + 1
'                   End If
                   
                    i = i + 1
                  
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
                
                 If controle = lnsap Or i = contador Then  ' verificar transferencia e criar new page
                 
                    varLn = 1
                    session.findById("wnd[0]").sendVKey 0
             
                    texto = session.findById("wnd[0]/sbar/pane[0]").Text
                 
                    Do Until texto = ""
            
                     lotep = session.findById("wnd[0]/usr/ctxtMSEG-CHARG").Text
                         If dadosBatimento(varLn, 7) <> "" Then
                         
                            session.findById("wnd[0]/usr/ctxtMSEG-UMCHA").Text = "000" & dadosBatimento(varLn, 7)
            
                            session.findById("wnd[0]").sendVKey 0
                            texto = session.findById("wnd[0]/sbar/pane[0]").Text
                            
                                Do Until varNewItem = True
                 
                                    If InStr(texto, "O lote recebedora é definido automaticamente") > 0 Then ' lote transferido , incrementar linha mamtriz
                                        'session.findbyid("wnd[0]").sendVKey 0
                                        'texto = ""
                                        'ccLinha = ccLinha + 1
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
                         
                         varLn = varLn + 1
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


