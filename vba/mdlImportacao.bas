Attribute VB_Name = "mdlImportacao"
'==============================================================================
' Importa a base de lançamentos (CSV ; exportado do sistema contábil) para a
' aba Lancamentos, sem duplicar IDs já existentes.
' Autor: Gabriel Antônio Ribeiro
'==============================================================================
Option Explicit

Public Sub ImportarLancamentos()
    Dim arq As Variant, wb As Workbook, origem As Worksheet, destino As Worksheet
    Dim ids As Object, i As Long, prox As Long, ultimaOrig As Long, novos As Long, c As Long

    arq = Application.GetOpenFilename("Arquivos CSV (*.csv),*.csv", , "Selecione o CSV de lançamentos")
    If VarType(arq) = vbBoolean Then Exit Sub          ' usuário cancelou

    Application.ScreenUpdating = False
    On Error GoTo Falha
    Set destino = ThisWorkbook.Worksheets("Lancamentos")

    Set ids = CreateObject("Scripting.Dictionary")
    For i = 2 To destino.Cells(destino.Rows.Count, 1).End(xlUp).Row
        ids(CStr(destino.Cells(i, 1).Value)) = True
    Next i

    Workbooks.OpenText Filename:=arq, Origin:=65001, DataType:=xlDelimited, Semicolon:=True, Local:=True
    Set wb = ActiveWorkbook
    Set origem = wb.Worksheets(1)
    ultimaOrig = origem.Cells(origem.Rows.Count, 1).End(xlUp).Row
    prox = destino.Cells(destino.Rows.Count, 1).End(xlUp).Row + 1

    For i = 2 To ultimaOrig
        If Not ids.Exists(CStr(origem.Cells(i, 1).Value)) Then
            For c = 1 To 6
                destino.Cells(prox, c).Value = origem.Cells(i, c).Value
            Next c
            destino.Cells(prox, 2).NumberFormat = "@"
            ids(CStr(origem.Cells(i, 1).Value)) = True
            prox = prox + 1: novos = novos + 1
        End If
    Next i

    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.Calculate
    MsgBox novos & " lançamento(s) importado(s). Duplicados foram ignorados.", vbInformation
    Exit Sub
Falha:
    Application.ScreenUpdating = True
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
End Sub
