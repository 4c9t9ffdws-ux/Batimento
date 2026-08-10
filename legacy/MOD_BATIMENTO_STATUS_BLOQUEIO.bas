Attribute VB_Name = "MOD_BATIMENTO_STATUS_BLOQUEIO"
Sub sb_Batimento_Status_BLOQUEIO()


    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim ln As Long
    Dim saldoNecessario As Double
    Dim saldoLivreLinhaAtual As Double
    Dim saldoBloqueadoLinhaAtual As Double
    Dim dlxBloqueadoLinhaAtual As Double
    Dim faltante As Double
    Dim sobrante As Double
    Dim linhaSaida As Long
    Dim lnAUX As Integer
    Dim cLN_matriz  As Integer
    Dim qtdItens    As Integer
    Dim transf      As Boolean
    Dim alteracoes  As New Collection
    Dim qtdAlteracao    As Integer
    Dim errorT          As Boolean

    qtdItens = 0
    qtdAlteracao = 0
    lastRow = ws_base.Cells(Rows.Count, "V").End(xlUp).Row
    ReDim dadosBatimento(1 To lastRow * 2, 1 To 8)
    
    transf = False
    errorT = False
     'CONEXÃO

    Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
    Set session = SAPCon.Children(0) 'Get the first session (window)
    
linhaSAP = 0

    ln = 6
    For i = 1 To lastRow - 5
        dadosBatimento(i, 1) = ws_base.Cells(ln, "s").Value ' MATERIAL
        dadosBatimento(i, 2) = ws_base.Cells(ln, "u").Value ' lote
        dadosBatimento(i, 3) = ws_base.Cells(ln, "V").Value ' SAP Livre
        dadosBatimento(i, 4) = ws_base.Cells(ln, "W").Value ' SAP Bloqueado
        dadosBatimento(i, 5) = ws_base.Cells(ln, "Z").Value ' DLX Bloqueado
        ln = ln + 1
        qtdItens = qtdItens + 1
    Next i



    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1B"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "344"

    If resposta2 = 1 Then
        session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7962"
    Else
    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7919"
    End If

    session.findById("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
    
     session.findById("wnd[0]").sendVKey 0
     
' --------

cLN_matriz = qtdItens + 1

    For i = 1 To qtdItens
    
        saldoLivreLinhaAtual = dadosBatimento(i, 3)
        saldoBloqueadoLinhaAtual = dadosBatimento(i, 4)
        dlxBloqueadoLinhaAtual = dadosBatimento(i, 5)
        
        saldoNecessario = dlxBloqueadoLinhaAtual - saldoBloqueadoLinhaAtual
        
        If saldoNecessario <= saldoLivreLinhaAtual Then
             'Saldo SAP é suficiente, mas deve ser limitado ao saldo livre atual
            dadosBatimento(i, 6) = saldoNecessario
            dadosBatimento(i, 7) = dadosBatimento(i, 2)
            
          Else
          
            dadosBatimento(i, 6) = saldoLivreLinhaAtual
            dadosBatimento(i, 7) = dadosBatimento(i, 2)
        End If
        
        

    Next i
    
  contador = 0
   'verificar quantos itens temos
   
   For ln = 1 To UBound(dadosBatimento, 1)
   
    If dadosBatimento(ln, 1) <> "" Then
    
        contador = contador + 1
        
    End If
   
   Next ln
   
linhaSAP = 0

' Inserir no SAP

   
        
    lnsap = session.findById("wnd[0]/usr/sub:SAPMM07M:0421").loopRowCount ' CONTAR Quantas linhas temos na tela sap
    
    
If lnsap >= contador Then ' tenho qtd suficiente para 1 pagina

    'loop para inserir itens
    
          For i = 1 To contador ' item 1 ate a quantidade de itens que esta na variavel contador
          
          
           'somente inserir valores contidos
           
               If dadosBatimento(i, 6) <> 0 Then
              
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[" & linhaSAP & ",7]").Text = dadosBatimento(i, 1)
                       
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[" & linhaSAP & ",26]").Text = dadosBatimento(i, 6)
                
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[" & linhaSAP & ",44]").Text = "CX"
                
                session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[" & linhaSAP & ",53]").Text = "000" & dadosBatimento(i, 7)
        
                linhaSAP = linhaSAP + 1
               End If
               
          Next i


i = 1
       session.findById("wnd[0]").sendVKey 0
             
      texto = session.findById("wnd[0]/sbar/pane[0]").Text
      
            Do Until texto = ""
            
            lotep = session.findById("wnd[0]/usr/ctxtMSEG-CHARG").Text
            'loteT = session.findbyid("wnd[0]/usr/ctxtMSEG-UMCHA").Text
            
            
            
                 session.findById("wnd[0]").sendVKey 0
                 texto = session.findById("wnd[0]/sbar/pane[0]").Text
                 
                    If InStr(texto, "0 peças") > 0 Then
                        session.findById("wnd[0]").sendVKey 0
                        texto = ""
                        
                    End If
                    
                     If InStr(texto, "LO Livre utilização não atingido") > 0 Then
                        errorT = True
                        session.findById("wnd[0]").maximize
                        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1B"
                        session.findById("wnd[0]").sendVKey 0
                        session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "344"
                        
                            If resposta2 = 1 Then
                                session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7962"
                            Else
                            session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7919"
                            End If
                        
                        session.findById("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
                        texto = ""
                        
                        eErros = eErros + V_aux
                    End If
                
                Loop
    

Else ' + de uma pagina


        linhaSAP = 0
        contador = 1 ' controle das linhas para saber hora da nova pagina
        validationLinha = 1
        
            For i = 1 To qtdItens ' faça da linha 6 ate a ultima linha
            
             
                 If dadosBatimento(i, 6) <> 0 Then
      
                    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[" & linhaSAP & ",7]").Text = dadosBatimento(i, 1)
                           
                    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[" & linhaSAP & ",26]").Text = dadosBatimento(i, 6)
                    
                    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[" & linhaSAP & ",44]").Text = "CX"
                    
                    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[" & linhaSAP & ",53]").Text = "000" & dadosBatimento(i, 7)
            
                    
                    
                     If contador = lnsap Then
                  
                        session.findById("wnd[0]").sendVKey 0
             
                        texto = session.findById("wnd[0]/sbar/pane[0]").Text
            
                            Do Until texto = ""
                    
                                session.findById("wnd[0]").sendVKey 0
                                texto = session.findById("wnd[0]/sbar/pane[0]").Text
                                
                                 If InStr(texto, "0 peças") > 0 Then
                                    session.findById("wnd[0]").sendVKey 0
                                    texto = ""
                        
                                 End If
                    
                                
                                If InStr(texto, "LO Livre utilização não atingido") > 0 Then
                                    errorT = True
                                    session.findById("wnd[0]").maximize
                                    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1B"
                                    session.findById("wnd[0]").sendVKey 0
                                    session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "344"
                                    
                                        If resposta2 = 1 Then
                                            session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7962"
                                        Else
                                        session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7919"
                                        End If
                                    
                                    session.findById("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
                                    texto = ""
                                    
                                End If
                                
                            Loop
                      linhaSAP = 0
                      contador = 1 ' controle das linhas para saber hora da nova pagina
                      'inserir new pagina
                      session.findById("wnd[0]/tbar[1]/btn[19]").press
                      validationLinha = 2 ' variavel de segurança para garantir que a linha nao seja incrementada incorreta
                  
                   End If
            
                            If validationLinha = 1 Then
                            
                              linhaSAP = linhaSAP + 1
                              contador = contador + 1
                              
                            End If
                                  
                    
       
                End If
               
               Next i
               
                
End If


'salvar

If errorT = False Then
session.findById("wnd[0]/tbar[0]/btn[11]").press
End If
contador = 0
End Sub



