Attribute VB_Name = "Módulo1"
Option Explicit
Function mensagens_sistema(indice) As String
    Dim MENSAGENS As Variant
    MENSAGENS = Array("Pedidos com Montagens", "Salvo com Sucesso!", "Confirmação de e-mail")
    
    mensagens_sistema = MENSAGENS(indice)
    
End Function

Sub ExportFile(control As IRibbonControl)
    'Desabilita alertas e mensagens.
    Application.DisplayAlerts = False
    
    'Variáveis de controle
    Dim arquivo As Workbook
    Dim este As Workbook
    Dim plan As String
    
    'Definições de variáveis
    plan = "MONTAGENS"
    Set este = Application.ThisWorkbook
    Set arquivo = Application.Workbooks.Add
    
    'exporta a planilha para um novo arquivo e salva
    este.Sheets(plan).Copy after:=arquivo.Sheets(1)
    arquivo.Sheets(1).Delete
    arquivo.SaveAs este.Path & "/Montagens", 51
    arquivo.Close
    
    'Chama a função que envia emails
    If MsgBox("Arquivo salvo com sucesso, deseja enviar por email?", vbYesNo, "Montagens") = vbYes Then
        EnviarEmail Application.ThisWorkbook.Path & "/Montagens.xlsx", Sheets("email").Range("B2"), Sheets("email").Range("B3"), Sheets("email").Range("B4"), _
            Sheets("email").Range("B5")
    End If
    'habilita os alertas e mensagens
    Application.DisplayAlerts = True
    
End Sub

Sub SettingMail(control As IRibbonControl)
    frmAtualizaEmail.Show
    
End Sub

Sub SendMail(control As IRibbonControl)
    EnviarEmail Application.ThisWorkbook.Path & "/Montagens.xlsx", Sheets("email").Range("B2"), Sheets("email").Range("B3"), Sheets("email").Range("B4"), _
        Sheets("email").Range("B5")
        
End Sub

Sub SettingUpdate(control As IRibbonControl)
    
    'Define as variáveis
    Dim historico As Workbook
    Dim limite As Integer
    Dim docPrecedente As Long
    Dim proc As Variant
    
    'abre o arquivo do histórico
    Set historico = Application.Workbooks.Open(Application.ThisWorkbook.Path & "/Histórico.xlsx")
    
    'verifica se o histórico já está atualizado
    docPrecedente = Application.ThisWorkbook.Sheets("MONTAGENS").Range("E2").Value
    
    On Error Resume Next:
        proc = Application.WorksheetFunction.VLookup(docPrecedente, historico.Sheets(1).Range("A:D"), 3, 0)
    On Error GoTo erro:
    
    If proc = "" Then
        'Caso o valor não seja encontrado, atualiza o histórico
        limite = historico.Sheets(1).Range("A100000").End(xlUp).Row + 1
        Application.ThisWorkbook.Sheets("MONTAGENS").Range("MONTAGENS[[Documento precedente]:[Cliente]]").Copy
        historico.Sheets(1).Range("A" & limite).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        historico.Sheets(1).Range("A1").Select
        historico.save
        historico.Close
        MsgBox "Histórico Atualizado com Sucesso!", vbExclamation, mensagens_sistema(0)
        
    Else
        historico.Close
        MsgBox "Histórico já está atualizado!!!", vbInformation, mensagens_sistema(0)
        
    End If
    
    Exit Sub
erro:
    
    MsgBox "Erro: " & Err.Number & vbNewLine & "Descrição: " & Err.Description, vbCritical, mensagens_sistema(0)
    
    
End Sub

Sub Infos(control As IRibbonControl)
    With Sheets("Informações de Atualização")
        .Visible = True
        .Select
    End With
    
End Sub

Sub bacHome()
    Sheets("Informações de Atualização").Visible = False
    Sheets("MONTAGENS").Select
    Range("A1").Select
    
End Sub
