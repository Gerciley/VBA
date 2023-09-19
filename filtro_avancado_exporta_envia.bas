Attribute VB_Name = "Módulo1"
Option Explicit
Global numero_base As Integer

Function filtro_avancado() As Integer
Attribute filtro_avancado.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim criterio As Range
    Set criterio = Sheets("ZSDR069_OTIF").Range("BE1:BE2")
    
    Sheets("ZSDR069_OTIF").Columns("A:Y").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=criterio, CopyToRange:=Range( _
        "RECOLHIMENTO!Area_de_extracao"), Unique:=False
    Range("A1").Select
    
    numero_base = Sheets("RECOLHIMENTO").Range("A1000000").End(xlUp).Row

End Function

Function exportar_arquivo(nome)
    Dim arquivo As Workbook
    
    Set arquivo = ThisWorkbook
    Sheets("RECOLHIMENTO").Copy
    Set arquivo = ActiveWorkbook
    
    arquivo.SaveAs Application.ThisWorkbook.Path & "/envios/" & nome & ".xlsx", 51
    arquivo.Close
End Function

Function EnviarEmail(anexo, dest, cc, assunto, msg, index)

    'Variáveis
    Dim F_anexo As Variant
    Dim F_destinatarios As String
    Dim F_cc As String
    Dim F_assunto As String
    Dim F_mensagem As String
    Dim objOutlook As Object
    Dim email As Object
    Dim control As Long
    
    'criando Objeto de email
    Set objOutlook = CreateObject("Outlook.Application")
    Set email = objOutlook.createitem(0)
    
    'Setando variáveis de tipo texto
    F_anexo = anexo
    F_destinatarios = dest
    F_assunto = assunto
    F_mensagem = msg
    F_cc = cc
    
    'Criando e estruturando email
    email.display
    email.to = F_destinatarios
    email.cc = F_cc
    email.Subject = F_assunto & Date
    email.htmlbody = F_mensagem & email.htmlbody
    
    'Anexa os arquivos
    For control = 0 To 1
        email.attachments.Add (Application.ThisWorkbook.Path & "/envios/" & F_anexo(control) & ".xlsx")
        
        If index = 1 Then Exit For
    
    Next control
    'Envia o Email
    email.send
    
    MsgBox "Email enviado com Sucesso", vbOKOnly, "Confirmação de e-mail"
End Function


Sub main()
    'Variáveis de controle exportação de arquivos
    Dim data As Date
    Dim parametros As Variant
    Dim nomes As Variant
    Dim nome_arquivo As String
    Dim index As Long
    Dim envia_email As VbMsgBoxResult
    
    'Variáveis para envio de e-mail
    Dim destinatarios As String
    Dim com_copia As String
    Dim assunto As String
    Dim mensagem As String
    
    destinatarios = Sheets("DADOS").Range("B2")
    com_copia = Sheets("DADOS").Range("B3")
    assunto = Sheets("DADOS").Range("B4")
    mensagem = Sheets("DADOS").Range("B5")
    
    parametros = Array("<>Z2", "Z2")
    data = Format(Date, "dd-MM-yyyy")
    nome_arquivo = Replace(data, "/", "-")
    nomes = Array(nome_arquivo & " RECOLHIMENTO", nome_arquivo & " MOSTRUARIO")
    
    
    For index = 0 To 1
        Sheets("ZSDR069_OTIF").Range("BE2").Value = parametros(index)
        filtro_avancado
        
        If numero_base > 1 Then
            exportar_arquivo nomes(index)
        Else
            Exit For
        End If
            
    Next index
    
    envia_email = MsgBox("Arquivos Exportados, gostaria de fazer o envio agora?", vbYesNo, "Enviar Recolhimento Somerlog")
    
    If envia_email = vbYes Then EnviarEmail nomes, destinatarios, com_copia, assunto, mensagem, 1
    
End Sub
