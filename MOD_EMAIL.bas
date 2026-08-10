Attribute VB_Name = "MOD_EMAIL"
 Sub e_mail_Batimento()
'
Windows("Batimento Operacional.xlsm").Activate

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False
    
Dim saudacao        As String
  
  
If Hour(Now) < 12 Then
    saudacao = "Bom dia,"
ElseIf Hour(Now) < 18 Then
    saudacao = "Boa tarde,"
Else
    saudacao = "Boa noite,"
End If


ws_dados.Select

Set obejto_outlook = CreateObject("Outlook.Application")
Set Email = obejto_outlook.CreateItem(0)

Dim strbody As String

'ativa tela para visualização
Email.Display
'
Email.To = "dl-pbftaticooperacionalcabreuva@pepsico.com" ' range para pegar contatos a enviar e-mail.
Email.CC = "TticoOperacional-Cabreva@pepsico.onmicrosoft.com" 'range para pegar contatos a enviar e-mail com copia
'Email.To = "Marcos.deoliveira1@pepsico.com" ' range para pegar contatos a enviar e-mail.
'Email.CC = "Marcos.deoliveira1@pepsico.com" 'range para pegar contatos a enviar e-mail com copia
Email.Subject = "REPORT_BATIMENTO_LOTES" & " - " & CENTRO  ' Assunto do E-mail.


'Texto em HTML

texto1 = "Segue dados do Batimento Realizado : " & CENTRO & "<br><br>"
texto2 = "Lotes Cadastrados no Aplicativo de Controle de Lotes :" & "<br><br>"
texto3 = "Linhas que Apresentaram Erros SAP :" & eErros & "<br><br>"
'texto6 = "<link>https://forms.gle/tGbgRcdu38E2Wk6u8 </link>"

qtdAplicativo = Application.WorksheetFunction.CountA(ws_Arquivos.Range("W:W"))

strbody = htmlBody
'strbody2 = htmlBody 'teste de concatenar duas fontes falho

strbody = "<font style=""font-family: Consolas; font-size: 9pt;""/font>"
'strbody2 = "<font style=""font-family: Montserrat; font-size: 9pt;font-weight: bold;""/font>" 'desativado temporariamente

If qtdAplicativo > 3 Then
'chamada dos texto / / primeira linha configura fonte e chama texto/ segunda a tabela/terceira assinatura
Email.htmlBody = strbody & saudacao & "<br><br>" & texto1 & texto3 _
 & RangetoHTML(ws_dados.Range("A1").CurrentRegion) _
& "<br><br>" & texto2 & RangetoHTML(ws_Arquivos.Range("w5").CurrentRegion) & Email.htmlBody


Else

Email.htmlBody = strbody & saudacao & "<br><br>" & texto1 & texto3 _
 & RangetoHTML(ws_dados.Range("A1").CurrentRegion) _
& "<br><br>" & Email.htmlBody

End If
'EMAIL.Attachments.Add ("C:\Users\09263374\Desktop\AjustePlanilhaEnviarEMAIL\passo a passo POD.mp4") 'campo pode ser utilizado para anexos

'Enviar e-mail

Email.send

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.EnableEvents = True
    

End Sub
 
Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function


