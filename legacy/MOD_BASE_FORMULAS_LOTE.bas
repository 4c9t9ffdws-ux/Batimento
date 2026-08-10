Attribute VB_Name = "MOD_BASE_FORMULAS_LOTE"
Sub formulas_da_base_LOTE()

If resposta3 = 1 Then
'robo
    ws_base.Range("B6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!I:I;ARQUIVOS!B:B);PROCX(A6;ARQUIVOS!BM:BM;ARQUIVOS!BG:BG))"
    ws_base.Range("C6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!I:I;ARQUIVOS!C:C);PROCX(A6;ARQUIVOS!BM:BM;ARQUIVOS!BH:BH))"
    ws_base.Range("D6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!I:I;ARQUIVOS!A:A);PROCX(A6;ARQUIVOS!BM:BM;ARQUIVOS!BI:BI))"
    
 Else
 'manual
     ws_base.Range("B6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!AO:AO;ARQUIVOS!AH:AH);PROCX(A6;ARQUIVOS!BM:BM;ARQUIVOS!BG:BG))"
    ws_base.Range("C6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!AO:AO;ARQUIVOS!AI:AI);PROCX(A6;ARQUIVOS!BM:BM;ARQUIVOS!BH:BH))"
    ws_base.Range("D6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!AO:AO;ARQUIVOS!AG:AG);PROCX(A6;ARQUIVOS!BM:BM;ARQUIVOS!BI:BI))"
 
 End If
    
    'zm
    ws_base.Range("E6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!BM:BM;ARQUIVOS!BJ:BJ);0)"
    ws_base.Range("F6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!BM:BM;ARQUIVOS!BK:BK);0)"
    ws_base.Range("G6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!BM:BM;ARQUIVOS!BL:BL);0)"
    
 
If resposta3 = 1 Then
'robo
    ws_base.Range("H6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!I:I;ARQUIVOS!F:F);0)"
    ws_base.Range("I6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!I:I;ARQUIVOS!G:G);0)"
    ws_base.Range("J6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!I:I;ARQUIVOS!H:H);0)"
    
Else
'manual
    ws_base.Range("H6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!AO:AO;ARQUIVOS!AL:AL);0)"
    ws_base.Range("I6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!AO:AO;ARQUIVOS!AM:AM);0)"
    ws_base.Range("J6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!AO:AO;ARQUIVOS!AN:AN);0)"

End If

    ws_base.Range("K6").Formula2Local = "=h6-e6"
    'ws_base.Range("L6").Formula2Local = "=SE(E(N6=""D"";H6=0);0;MÍNIMO(F6- I6; H6))"
    'ws_base.Range("M6").Formula2Local = "=I6-(E6+F6)"
   ' ws_base.Range("N6").Formula2Local = "=SE(I6-F6=0;0;SE(E(I6-F6<0;H6<>0);""D"";SE(E(I6-F6<0;H6=0);0;""B"")))"
    ws_base.Range("O6").Formula2Local = "=PROCX(B6;B:B;D:D)"
    
 If resposta3 = 1 Then
 'robo
    ws_base.Range("P6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!I:I;ARQUIVOS!J:J);""OK"")"
    
 Else
 'manual
    ws_base.Range("P6").Formula2Local = "=SEERRO(PROCX(A6;ARQUIVOS!AO:AO;ARQUIVOS!AP:AP);""OK"")"
 
 End If
 
    ' Contar a última linha preenchida na coluna A
    ultimaLinha = ws_base.Cells(ws_base.Rows.Count, "A").End(xlUp).Row
    
    ' Arrastar as fórmulas da linha 6 até a última linha preenchida
    ws_base.Range("B6:P6").AutoFill Destination:=ws_base.Range("B6:P" & ultimaLinha)
    
    ' Transformar fórmulas em valores (de B até P) para todas as linhas preenchidas
    ws_base.Range("B6:P" & ultimaLinha).Value = ws_base.Range("B6:P" & ultimaLinha).Value


End Sub

