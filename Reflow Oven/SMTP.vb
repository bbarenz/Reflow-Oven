Imports System.Net.Mail
Imports System.Collections.Specialized
Imports System.Net
Imports System.Threading


Module SMTP
    '======================================================================
    ' Declare Others 

    '======================================================================

    Public Sub SendFeedback()



        Dim mail As New MailMessage()

        Dim SmtpServer As New SmtpClient()
        SmtpServer.Credentials = New Net.NetworkCredential("bei.errorlog@gmail.com", "errorlog")
        SmtpServer.Port = 587
        SmtpServer.Host = "smtp.gmail.com"
        SmtpServer.EnableSsl = True

        mail = New MailMessage()

        Try
            mail.From = New MailAddress("bei.errorlog@gmail.com", "Reflow Oven Console Feedback", System.Text.Encoding.UTF8)

            mail.To.Add("bei.errorlog@gmail.com")
            mail.Subject = "Reflow Oven Console Feedback"

            ' build the message to send 
            '---------------------------------------------------------------------------
            mail.Body = "Subject: " & Form1.txt_OtherSubjectText.Text & vbNewLine & Form1.rtxt_FeedbackMsg.Text
            '---------------------------------------------------------------------------



            SmtpServer.Send(mail)
        Catch ex As Exception
            '  MsgBox(ex.ToString())
        End Try

    End Sub

End Module
