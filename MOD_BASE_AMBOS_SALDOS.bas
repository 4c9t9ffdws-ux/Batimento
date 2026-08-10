Attribute VB_Name = "MOD_BASE_AMBOS_SALDOS"
Sub sb_base_status_lote()


 ws_Arquivos.Select
  'iNSERIR UMA LINHA COM QUARENTENA


If ActiveSheet.FilterMode Then
    
    ActiveSheet.ShowAllData     ' Limpa todos os filtros da planilha
    
    End If
    
  
    
' verificar se e manual ou robo

If resposta3 = 1 Then 'robo

ws_Arquivos.Range("SALDO_ROBO[[#Headers],[Número Lote]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
lnarq = ws_Arquivos.Cells(Rows.Count, "A").End(xlUp).Row + 1
   
'adicionar a linha que vai garantir que tenha quarentena
'mod extreme go horse -  ideia do victor

ws_Arquivos.Range("A" & lnarq).Value = "9999999"
ws_Arquivos.Range("B" & lnarq).Value = "9999999999"
ws_Arquivos.Range("C" & lnarq).Value = "GENERICO"
ws_Arquivos.Range("D" & lnarq).Value = "1"
ws_Arquivos.Range("E" & lnarq).Value = "Quarentena"

'atualizar a dinamica do saldo do robo

ActiveSheet.PivotTables("Tabela dinâmica1").PivotCache.Refresh

Else ' manual

Range("SALDO_MANUAL[[#Headers],[LOTE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

lnarq = ws_Arquivos.Cells(Rows.Count, "Ag").End(xlUp).Row + 1


ws_Arquivos.Range("AG" & lnarq).Value = "9999999"
ws_Arquivos.Range("ah" & lnarq).Value = "9999999999"
ws_Arquivos.Range("AI" & lnarq).Value = "GENERICO"
ws_Arquivos.Range("AJ" & lnarq).Value = "1"
ws_Arquivos.Range("AK" & lnarq).Value = "Quarentena"

'ATUALIZAR dinamica do saldo manual

ActiveSheet.PivotTables("Tabela dinâmica2").PivotCache.Refresh
End If

'atualizar a tabela que consta os dados do aplicativo de lote

Range("WHS_BATIMENTO_LOTES[[#Headers],[LOTE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    
 'zm atualizar
 
  Range("SAP_ZM[[#Headers],[ Material          ]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
  
'=============================================================================================================
    ' Verifica qual opção foi selecionada e qual o caminho a seguir
               

 If resposta3 = 1 Then ' robo
 
'CONTAR LINHAS

LnTBL = ws_Arquivos.Cells(Rows.Count, "A").End(xlUp).Row ' contar linhas saldo do rolo
'lnValidacao = ws_Arquivos.Cells(Rows.Count, "bp").End(xlUp).Row
 
'=============================================================================================================
  'VALIDAR SE O CENTRO FOI INFORMADO CORRETO
  
  If resposta2 = 1 And LnTBL < 8000 Or resposta2 = 2 And LnTBL > 8000 Then
  
   MsgBox "Saldo Selecionado Incorretamente!!" & vbCrLf & cCentro & " com total de: " & LnTBL & " - Linhas", vbCritical, "Batimento de lote"
   validationSaldo = True
   Exit Sub
  End If
'=============================================================================================================
  
'copiar e mover MATERIAL

ws_Arquivos.Range("B2:B" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bp2").PasteSpecial
 
'=============================================================================================================
  
  
'copiar e mover LOTE

ws_Arquivos.Range("A2:A" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bQ2").PasteSpecial
 
'=============================================================================================================
'copiar e mover CHAVE

ws_Arquivos.Range("I2:I" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bR2").Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
'=============================================================================================================
       
'colar a zm

LnTBL = ws_Arquivos.Cells(Rows.Count, "BG").End(xlUp).Row ' contar linhas saldo do rolo
lnValidacao = ws_Arquivos.Cells(Rows.Count, "bp").End(xlUp).Row + 1
'copiar e mover MATERIAL

ws_Arquivos.Range("BG6:BG" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bp" & lnValidacao).PasteSpecial
 
'=============================================================================================================
  
'copiar e mover LOTE

ws_Arquivos.Range("BI6:BI" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bQ" & lnValidacao).PasteSpecial
 
'=============================================================================================================
     
     
'copiar e mover CHAVE

ws_Arquivos.Range("BM6:BM" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bR" & lnValidacao).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
'=============================================================================================================
   
 'RETIRAR AS DUPLICADAS
 
 lnValidacao = ws_Arquivos.Cells(Rows.Count, "bp").End(xlUp).Row
 
 ActiveSheet.Range("BP2:BR" & lnValidacao).RemoveDuplicates Columns:=3, _
    Header:=xlYes
' ActiveSheet.Range("BP2:BR" & lnValidacao).RemoveDuplicates Columns:=Array(1), _
        Header:=xlYes
'=============================================================================================================
 ' LEVAR A CHAVE PARA BASE
   
lnValidacao = ws_Arquivos.Cells(Rows.Count, "bp").End(xlUp).Row

ws_Arquivos.Range("br2:br" & lnValidacao).Select

Selection.Copy

'=============================================================================================================
 ' levar para base
    
 ws_base.Range("A6").PasteSpecial

Else ' manual ---------- <
 
 
LnTBL = ws_Arquivos.Cells(Rows.Count, "AG").End(xlUp).Row ' contar linhas saldo do rolo

 
 '=============================================================================================================
  'VALIDAR SE O CENTRO FOI INFORMADO CORRETO
  
  If resposta2 = 1 And LnTBL < 8000 Or resposta2 = 2 And LnTBL > 8000 Then
  
   MsgBox "Saldo Selecionado Incorretamente!!" & vbCrLf & cCentro & " com total de: " & LnTBL & " - Linhas", vbCritical, "Batimento de lote"
   validationSaldo = True
   Exit Sub
  End If
'==========

'copiar e mover MATERIAL

ws_Arquivos.Range("AH2:AH" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bp2").PasteSpecial
 
'=============================================================================================================
  
'copiar e mover LOTE

ws_Arquivos.Range("AG2:AG" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bQ2").PasteSpecial
 
'=============================================================================================================
     
     
'copiar e mover CHAVE

ws_Arquivos.Range("AO2:AO" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bR2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
'=============================================================================================================
         
'colar a zm

LnTBL = ws_Arquivos.Cells(Rows.Count, "BG").End(xlUp).Row ' contar linhas saldo do rolo
lnValidacao = ws_Arquivos.Cells(Rows.Count, "bp").End(xlUp).Row + 1
'copiar e mover MATERIAL

ws_Arquivos.Range("BG6:BG" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bp" & lnValidacao).PasteSpecial
 
'=============================================================================================================
  
'copiar e mover LOTE

ws_Arquivos.Range("BI6:BI" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bQ" & lnValidacao).PasteSpecial
 
'=============================================================================================================
     
     
'copiar e mover CHAVE

ws_Arquivos.Range("BM6:BM" & LnTBL).Select

Selection.Copy

ws_Arquivos.Range("bR" & lnValidacao).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
'=============================================================================================================
 
  'RETIRAR AS DUPLICADAS
 
 lnValidacao = ws_Arquivos.Cells(Rows.Count, "bp").End(xlUp).Row
 
 ActiveSheet.Range("BP2:BR" & lnValidacao).RemoveDuplicates Columns:=3, _
    Header:=xlYes
' ActiveSheet.Range("BP2:BR" & lnValidacao).RemoveDuplicates Columns:=Array(1), _
        Header:=xlYes
'=============================================================================================================
 ' LEVAR A CHAVE PARA BASE
   
lnValidacao = ws_Arquivos.Cells(Rows.Count, "bp").End(xlUp).Row

ws_Arquivos.Range("br2:br" & lnValidacao).Select

Selection.Copy

'=============================================================================================================
 ' levar para base
    
 ws_base.Range("A6").PasteSpecial
 End If '< -------FINAL IF
 


End Sub

