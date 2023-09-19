Attribute VB_Name = "Módulo2"
Option Explicit
'Função pública que envia email
Public Function EnviarEmail(anexo, dest, cc, assunto, msg)

'Variáveis
Dim F_anexo As String
Dim F_destinatarios As String
Dim F_cc As String
Dim F_assunto As String
Dim F_mensagem As String
Dim objOutlook As Object
Dim email As Object

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
email.Subject = F_assunto
email.htmlbody = F_mensagem & email.htmlbody

'Verifica se contem anexo, caso tenha anexa o documento do email
If Len(F_anexo) <> 0 Then email.attachments.Add (F_anexo)

'Envia o Email
'Stop
email.send

MsgBox "Email enviado com Sucesso", vbOKOnly, mensagens_sistema(2)
End Function
