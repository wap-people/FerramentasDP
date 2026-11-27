Attribute VB_Name = "TelegramaAssinatura"
Option Explicit

' ==========================================================
'  TELEGRAMA DE ABANDONO - rascunho no Outlook com a assinatura padrao
'  (mesma logica da macro do Checklist de Desligamento)
'
'  1. Na ferramenta do telegrama, clique em "Gerar e-mail com PDF anexo"
'     (baixa "Telegrama - NOME.eml" na pasta Downloads).
'  2. Rode TelegramaComAssinatura (funciona no Outlook ou no Excel).
'
'  Pega o .eml mais recente, cria um e-mail novo pelo Outlook
'  (GetInspector carrega a assinatura padrao), poe o texto ANTES da
'  assinatura, copia Para/Cc/Cco, assunto e o PDF, salva em Rascunhos
'  e apaga o .eml baixado (tem dados pessoais do colaborador).
'  Sem acentos no codigo: os textos usam ChrW - nao corrompem.
' ==========================================================

Public Sub TelegramaComAssinatura()
    Const PREFIXO As String = "Telegrama - "
    Dim pasta As String, arq As String, maisNovo As String, dtNovo As Date
    Dim olApp As Object, origem As Object, olMail As Object, olInsp As Object
    Dim r As Object, nr As Object, a As Object, tmp As String
    Dim html As String, texto As String, resto As String, re As Object
    Dim posBody As Long, posFecha As Long

    ' ---- .eml mais recente da pasta Downloads ----
    pasta = Environ$("USERPROFILE") & "\Downloads\"
    arq = Dir$(pasta & PREFIXO & "*.eml")
    Do While arq <> ""
        If FileDateTime(pasta & arq) > dtNovo Then
            dtNovo = FileDateTime(pasta & arq)
            maisNovo = arq
        End If
        arq = Dir$()
    Loop
    If maisNovo = "" Then
        MsgBox "N" & ChrW(227) & "o encontrei nenhum arquivo '" & PREFIXO & "*.eml' na pasta Downloads." & vbCrLf & _
               "Gere o e-mail na ferramenta do telegrama e rode a macro de novo.", vbExclamation, "Telegrama"
        Exit Sub
    End If

    On Error GoTo FalhaOutlook
    Set olApp = CreateObject("Outlook.Application")
    Set origem = olApp.Session.OpenSharedItem(pasta & maisNovo)

    ' ---- e-mail novo: o inspector carrega a assinatura padrao ----
    Set olMail = olApp.CreateItem(0)
    Set olInsp = olMail.GetInspector

    For Each r In origem.Recipients
        Set nr = olMail.Recipients.Add(r.Address)
        nr.Type = r.Type                         ' 1 Para | 2 Cc | 3 Cco
    Next r
    olMail.Recipients.ResolveAll
    olMail.Subject = origem.Subject

    ' ---- texto do telegrama antes da assinatura ----
    texto = DentroDoBody(origem.HTMLBody)
    html = olMail.HTMLBody
    posBody = InStr(1, html, "<body", vbTextCompare)
    If posBody > 0 Then
        posFecha = InStr(posBody, html, ">")
        resto = Mid$(html, posFecha + 1)
        ' remove a linha em branco que o Outlook coloca antes da assinatura
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True: re.Global = False
        re.Pattern = "^(\s*<div[^>]*>)?\s*<p[^>]*>\s*(<o:p>)?(\s|&nbsp;)*(</o:p>)?\s*</p>"
        resto = re.Replace(resto, "$1")
        olMail.HTMLBody = Left$(html, posFecha) & texto & resto
    Else
        olMail.HTMLBody = texto & html
    End If

    ' ---- anexo (PDF do telegrama) ----
    For Each a In origem.Attachments
        tmp = Environ$("TEMP") & "\" & a.FileName
        a.SaveAsFile tmp
        olMail.Attachments.Add tmp
        Kill tmp
    Next a

    olMail.Save                                  ' pasta Rascunhos
    olInsp.Close 0
    origem.Close 1                               ' descarta o .eml aberto
    Set origem = Nothing
    Kill pasta & maisNovo                        ' apaga o .eml baixado

    MsgBox "Rascunho criado na pasta Rascunhos do Outlook!" & vbCrLf & olMail.Subject, vbInformation, "Telegrama"
    Set olInsp = Nothing: Set olMail = Nothing: Set olApp = Nothing
    Exit Sub

FalhaOutlook:
    MsgBox "N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel montar o rascunho: " & Err.Description, vbCritical, "Telegrama"
End Sub

' conteudo de dentro do <body> do e-mail gerado pela ferramenta
Private Function DentroDoBody(ByVal html As String) As String
    Dim i As Long, j As Long, k As Long
    i = InStr(1, html, "<body", vbTextCompare)
    If i = 0 Then DentroDoBody = html: Exit Function
    j = InStr(i, html, ">")
    k = InStr(j, html, "</body>", vbTextCompare)
    If k = 0 Then k = Len(html) + 1
    DentroDoBody = Mid$(html, j + 1, k - j - 1)
End Function
