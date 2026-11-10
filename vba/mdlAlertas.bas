Attribute VB_Name = "mdlAlertas"
'==============================================================================
' Controle de despesas recorrentes - envio de alertas de pendência
' Autor: Gabriel Antônio Ribeiro
'
' Agrupa as linhas PENDENTE da aba Recorrentes por destinatário e monta UM e-mail
' por destinatário no Outlook. Por segurança, o padrão é EXIBIR (abre o e-mail
' para revisão); só envia direto se Parametros!B4 = "ENVIAR".
' Usa late binding: não exige referência à biblioteca do Outlook.
'==============================================================================
Option Explicit

Private Const COL_FORNECEDOR As Long = 1
Private Const COL_PA As Long = 2
Private Const COL_TIPO As Long = 3
Private Const COL_STATUS As Long = 7
Private Const COL_DESTINO As Long = 8
Private Const COL_EMAIL As Long = 9

Public Sub GerarAlertasPendencias()
    Dim ws As Worksheet, grupos As Object, ultima As Long, i As Long
    Dim chave As Variant, competencia As String, modo As String, assinatura As String
    Dim enviados As Long

    On Error GoTo Falha
    Set ws = ThisWorkbook.Worksheets("Recorrentes")
    Application.Calculate

    competencia = CStr(ThisWorkbook.Worksheets("Parametros").Range("B3").Value)
    modo = UCase$(Trim$(CStr(ThisWorkbook.Worksheets("Parametros").Range("B4").Value)))
    assinatura = CStr(ThisWorkbook.Worksheets("Parametros").Range("B5").Value)
    If modo <> "ENVIAR" Then modo = "EXIBIR"

    Set grupos = CreateObject("Scripting.Dictionary")
    ultima = ws.Cells(ws.Rows.Count, COL_FORNECEDOR).End(xlUp).Row

    For i = 2 To ultima
        If ws.Cells(i, COL_STATUS).Value = "PENDENTE" Then
            chave = CStr(ws.Cells(i, COL_EMAIL).Value)
            If InStr(1, chave, "@") = 0 Then
                MsgBox "Linha " & i & ": destinatário '" & ws.Cells(i, COL_DESTINO).Value & _
                       "' sem e-mail na aba Destinatarios. Corrija antes de continuar.", vbExclamation
                Exit Sub
            End If
            If Not grupos.Exists(chave) Then grupos.Add chave, ""
            grupos(chave) = grupos(chave) & "<tr><td>" & ws.Cells(i, COL_FORNECEDOR).Value & "</td><td>" & _
                            ws.Cells(i, COL_PA).Value & "</td><td>" & ws.Cells(i, COL_TIPO).Value & "</td></tr>"
        End If
    Next i

    If grupos.Count = 0 Then
        MsgBox "Nenhuma pendência na competência " & competencia & ".", vbInformation
        Exit Sub
    End If

    If modo = "ENVIAR" Then
        If MsgBox("Serão ENVIADOS " & grupos.Count & " e-mails sem revisão. Confirma?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    For Each chave In grupos.Keys
        CriarEmail CStr(chave), competencia, CStr(grupos(chave)), assinatura, modo
        enviados = enviados + 1
    Next chave

    MsgBox enviados & " e-mail(s) " & IIf(modo = "ENVIAR", "enviado(s).", "aberto(s) para revisão."), vbInformation
    Exit Sub
Falha:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Sub CriarEmail(ByVal para As String, ByVal competencia As String, ByVal linhas As String, _
                       ByVal assinatura As String, ByVal modo As String)
    Dim ol As Object, msg As Object
    Set ol = CreateObject("Outlook.Application")
    Set msg = ol.CreateItem(0)              ' 0 = olMailItem
    With msg
        .To = para
        .Subject = "Pendência de documentos fiscais - competência " & competencia
        .HTMLBody = "<p>Olá,</p><p>Ainda não recebemos os documentos abaixo referentes à competência <b>" & _
                    competencia & "</b>. Pedimos o envio para que o pagamento não atrase:</p>" & _
                    "<table border='1' cellpadding='4' style='border-collapse:collapse;font-family:Calibri'>" & _
                    "<tr style='background:#003641;color:#fff'><th>Fornecedor</th><th>PA</th><th>Despesa</th></tr>" & _
                    linhas & "</table><p>Atenciosamente,<br>" & assinatura & "</p>"
        If modo = "ENVIAR" Then .Send Else .Display
    End With
End Sub
