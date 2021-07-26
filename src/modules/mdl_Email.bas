Attribute VB_Name = "mdl_Email"
Option Explicit
Const cSmtpAuthenticateOutlook = 1
Const cSmtpServerOutlook = "smtp-mail.outlook.com"
Const cSendUsingOutlook = 2
Const cSmtpServerPortOutlook = 25

Sub enviarEmail(pDados() As Variant)

Dim iMsg As Object
Dim iConf As Object
Dim tabela As String
Dim Flds As Variant
Dim i As Integer
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
    iConf.Load -1    ' CDO Source Defaults
    
    Set Flds = iConf.Fields
    
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cSmtpAuthenticateOutlook
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "seu_email@seu_email"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "sua_senha"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = cSmtpServerOutlook
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cSendUsingOutlook
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = cSmtpServerPortOutlook
        .Update
    End With

        tabela = tabela & styleCss
        tabela = tabela & "<div class=""center fundo full-height"">"
        tabela = tabela & header
        
        tabela = tabela & "<div class=""tabela"">"
        
        tabela = tabela & "<table align=""center"" border=""1"" cellpadding=""4"" cellspacing=""1""> "
        tabela = tabela & "<tbody>"
        
        tabela = tabela & "<thead align=""center"">"
        tabela = tabela & "<td>CLIENTE</td>"
        tabela = tabela & "<td>VALOR</td>"
        tabela = tabela & "<td>VENCIMENTO</td>"
        tabela = tabela & "<td>N° FATURA</td>"
        tabela = tabela & "</thead>"
    
    For i = 0 To UBound(pDados, 2)
        tabela = tabela & "<tr>"
        tabela = tabela & "<td>" & pDados(0, i) & "</td>"
        tabela = tabela & "<td>" & pDados(1, i) & "</td>"
        tabela = tabela & "<td>" & pDados(2, i) & "</td>"
        tabela = tabela & "<td>" & pDados(3, i) & "</td>"
        tabela = tabela & "</tr>"
    Next i
    
        tabela = tabela & "</tbody>"
        tabela = tabela & "</table>"
        tabela = tabela & "</div>"
        tabela = tabela & footer
    
    With iMsg
        Set .Configuration = iConf
        .To = "email_destinatario@email_destinatario"
        .CC = ""
        .BCC = ""
        .From = "seu_email@seu_email"
        .Subject = "Elemax Elevadores - FATURAS EM ABERTO"
        .HTMLBody = tabela 'para mensagem em texto, usar TextBody
        '.AddAttachment ("c:\fatura.txt") 'envia anexo caso precisar, foi testado e funcionou
        .Send
    End With
    
    Call salvaHistoricoDeEnvioDeEmails(pDados)

End Sub



Function styleCss() As String

styleCss = "<style> tr:nth-child(even) { background-color: #f2f2f2; } thead { background-color: black; color: white; font-weight: bold; } body .tabela { overflow-x: auto; } .center { margin: auto; width: 50%; padding: 10px; } .fundo { background-color: rgb(245, 245, 245); border: 1px solid #ccc !important; border-radius: 1%; font-family: Arial, Helvetica, sans-serif; } .full-height { height: 90%; } .footer { margin: 0px; text-align: center; font-weight: bold; } </style>"


End Function

Function footer() As String

    footer = "<p align=""center"">Temos certeza que o atraso se deve a problemas circunstanciais e que chegaremos logo a um acordo de negociação.</p> <p align=""center"">Quaisquer dúvidas estamos a inteira disposição. Você pode responder este e-mail ou entrar em contato no telefone abaixo em horário comercial.</p> <p align=""center"">Caso o pagamento tenha sido efetuado, por favor, desconsidere esta mensagem.</p> <p class=""footer"">Atenciosamente,</p> <p class=""footer"">Elemax Elevadores Ltda.-ME</p> <p class=""footer"">Contato: (13) 3495-8950</p>"

End Function

Function header() As String

    header = "<hr/><h2 align= ""center"">SERVIÇO AUTOMÁTICO DE COBRANÇA</h2> <hr /> <p align= ""center"">Prezados(as) senhores(as), constam em nossos registros o não pagamento das faturas listadas abaixo. O contato é para que possam regularizar a situação financeira;</p>"

End Function

Sub salvaHistoricoDeEnvioDeEmails(pDados() As Variant)

Dim ultimaLinha As Long
Dim i As Long

    With wsEmailsEnviados
    
        For i = LBound(pDados, 2) To UBound(pDados, 2)
            
            ultimaLinha = mdl_Util.ultimaLinhaEmBranco(wsEmailsEnviados)
            
            .Cells(ultimaLinha, 1) = mdl_Util.geraId(wsEmailsEnviados)
            .Cells(ultimaLinha, 2) = Date
            .Cells(ultimaLinha, 3) = pDados(0, i)
            .Cells(ultimaLinha, 4) = pDados(1, i)
            .Cells(ultimaLinha, 5) = CDate(pDados(2, i))
            .Cells(ultimaLinha, 6) = pDados(3, i)
            .Cells(ultimaLinha, 7) = pDados(4, i)
            .Cells(ultimaLinha, 8) = pDados(5, i)
            
        Next i
        
    End With
    

End Sub

