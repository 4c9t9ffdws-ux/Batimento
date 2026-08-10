Attribute VB_Name = "MOD_BATIMENTO_STATUS_DESB"
Sub sb_Batimento_Status_desbloqueio()


 
    Dim lastRow As Long, i As Long
    Dim sapLivre As Long, sapBloqueado As Long, sapQuarentena As Long, qtdCX As Long
    Dim dlxLivre As Long, dlxBloqueado As Long, dlxQuarentena As Long, linhaSAP As Integer, validationLinha As Integer
    Dim material As String, lote As String
    Dim errorT  As Boolean
 
 
 'CONEXÃO

    Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
    Set session = SAPCon.Children(0) 'Get the first session (window)
  errorT = False
    
  session.findById("wnd[0]").sendVKey 0
  
    ' Última linha da tabela
    lastRow = ws_base.Cells(Rows.Count, "s").End(xlUp).Row
    
  
      
    lnsap = session.findById("wnd[0]/usr/sub:SAPMM07M:0421").loopRowCount ' CONTAR Quantas linhas temos na tela sap
    
    
    
    ' Loop através de cada linha da tabela, começando na linha 2 (assumindo que a primeira linha são cabeçalhos)
   
    
   'verificação de qtd linhas Vs quatidade material temos
   
        If lnsap >= lastRow - 5 Then ' tenho qtd suficiente para 1 pagina
        
        linhaSAP = 0
            For i = 6 To lastRow ' i = 6 porque as informaçoes começam na linha 6
                
               'alimentar variaveis
               material = ws_base.Range("S" & i).Value
               qtdCX = ws_base.Range("AC" & i).Value
               lote = "000" & ws_base.Range("U" & i).Value
               sapLivre = ws_base.Range("v" & i).Value
               dlxLivre = ws_base.Range("y" & i).Value
               
               
               If qtdCX < 0 Then ' SE o numero estiver negativo , mudar para positivo
               
                qtdCX = qtdCX * -1
                
               End If
               
    ' validar aqui ----------------------  desbloqueio ativando , porem sem saldo no dlx
    
               If qtdCX > dlxLivre Then ' verificaçao se temos a quantidade no SAP
                    
                     qtdCX = dlxLivre
                
               End If
 
               
               
               
               session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[" & linhaSAP & ",7]").Text = material
               
               session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[" & linhaSAP & ",26]").Text = qtdCX
               
               session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[" & linhaSAP & ",44]").Text = "CX"
               
               session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[" & linhaSAP & ",53]").Text = lote
               
               
               linhaSAP = linhaSAP + 1
             Next i
             
             session.findById("wnd[0]").sendVKey 0
             
             texto = session.findById("wnd[0]/sbar/pane[0]").Text
             
           Do Until texto = ""
           
           If InStr(texto, "LO Bloqueado não atingido") > 0 Or InStr(texto, "não existe") > 0 Then
                        errorT = True
                        session.findById("wnd[0]").maximize
                        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1B"
                        session.findById("wnd[0]").sendVKey 0
                        session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "343"
                        
                            If resposta2 = 1 Then
                                session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7962"
                            Else
                            session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "7919"
                            End If
                                
                                'session.findbyid("wnd[0]/usr/ctxtRM07M-LGORT").Text = "IM01"
                                texto = ""
                                eErros = eErros + V_aux
                                
                            End If
           If errorT = False Then
                            
            session.findById("wnd[0]").sendVKey 0
            texto = session.findById("wnd[0]/sbar/pane[0]").Text
            
          End If
          
           Loop
             
    ' adicionar botao de salvar
    'desbloquio concluido
    
             
             
        Else ' qtd itens e maior que a quantidade linhas
        
        linhaSAP = 0
        contador = 1 ' controle das linhas para saber hora da nova pagina
            For i = 6 To lastRow ' faça da linha 6 ate a ultima linha
            
             'alimentar variaveis
               material = ws_base.Range("S" & i).Value
               qtdCX = ws_base.Range("AE" & i).Value
               lote = "000" & ws_base.Range("U" & i).Value
               sapLivre = ws_base.Range("v" & i).Value
               validationLinha = 1 ' variavel de segurança para garantir que a linha nao seja incrementada incorreta
               
                If qtdCX < 0 Then ' SE o numero estiver negativo , mudar para positivo
               
                    qtdCX = qtdCX * -1
                
                 End If
                 
               
                 If qtdCX > sapLivre Then ' verificaçao se temos a quantidade no SAP
                
                    qtdCX = sapLivre
                 
                 End If
                 
                  
               
               
                    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[" & linhaSAP & ",7]").Text = material
                    
                    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[" & linhaSAP & ",26]").Text = qtdCX
                    
                    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[" & linhaSAP & ",44]").Text = "CX"
                    
                    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[" & linhaSAP & ",53]").Text = lote
               
               
                  If contador = lnsap Then
                  
                    session.findById("wnd[0]").sendVKey 0
             
                    texto = session.findById("wnd[0]/sbar/pane[0]").Text
            
                            Do Until texto = ""
                    
                                session.findById("wnd[0]").sendVKey 0
                                texto = session.findById("wnd[0]/sbar/pane[0]").Text
                                
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
          
            Next i
        
     
        
        End If
    
If errorT = False Then
session.findById("wnd[0]/tbar[0]/btn[11]").press
End If
End Sub
