Imports System.IO
Imports System.Net.Mail

Public Class mailSender
    Public Sub hsSendConnectionLost2Clevergy(hs As SCP.BLL.HeatingSystem)
        Dim moduleName As String = "hsConnectionLost.html"
        Dim mailContent As String = loadModule(moduleName)
        mailContent = Replace(mailContent, "@!@hsId@!@", hs.hsId)
        mailContent = Replace(mailContent, "@!@Descr@!@", hs.Descr)
        mailContent = Replace(mailContent, "@!@lastRec@!@", FormatDateTime(hs.lastRec, DateFormat.GeneralDate))
        Dim DesImpianto As String = String.Empty
        Dim _i As SCP.BLL.Impianti = SCP.BLL.Impianti.Read(hs.IdImpianto)
        If Not _i Is Nothing Then
            DesImpianto = _i.DesImpianto
            _i = Nothing
        End If
        mailContent = Replace(mailContent, "@!@DesImpianto@!@", DesImpianto)

        Dim Message As New System.Net.Mail.MailMessage
        Message.From = New MailAddress("noReply@clevergy.it")
        Message.To.Add("sistemi@clevergy.it")
        Message.To.Add("andrea.ravasio@clevergy.it")
        Message.To.Add("marco.fagnano@clevergy.it")
        Message.To.Add("andrea.laini@clevergy.it")


        Dim _list As List(Of SCP.BLL.hs_UserAlert) = SCP.BLL.hs_UserAlert.List(hs.hsId)
        If Not _list Is Nothing Then
            For Each _obj As SCP.BLL.hs_UserAlert In _list
                If Not String.IsNullOrEmpty(_obj.emailaddress) Then
                    Message.To.Add(_obj.emailaddress)
                End If
            Next
            _list = Nothing
        End If

        Message.Subject = "(ALLARME) Connessione persa."
        Message.Body = mailContent
        Message.IsBodyHtml = True

        ' You won't need the calls to Message.Fields.Add()

        ' Replace SmtpMail.SmtpServer = server with the following:
        Dim smtp As New SmtpClient("mail.innowatio.it")

        'Dim smtp As New SmtpClient("smtp.mandrillapp.com")
        'smtp.Port = 587
        'smtp.EnableSsl = True
        'smtp.Credentials = New System.Net.NetworkCredential("clv.emcs.00@gmail.com", "gr4SmRnLxNn6hRFyW8lq-Q")
        Try
            smtp.Send(Message)

        Catch ex As Exception

        End Try
    End Sub

    Public Sub hsSendConnectionRestored2Clevergy(hs As SCP.BLL.HeatingSystem)
        Dim moduleName As String = "hsConnectionRestored.html"
        Dim mailContent As String = loadModule(moduleName)
        mailContent = Replace(mailContent, "@!@hsId@!@", hs.hsId)
        mailContent = Replace(mailContent, "@!@Descr@!@", hs.Descr)
        mailContent = Replace(mailContent, "@!@lastRec@!@", FormatDateTime(hs.lastRec, DateFormat.GeneralDate))
        Dim DesImpianto As String = String.Empty
        Dim _i As SCP.BLL.Impianti = SCP.BLL.Impianti.Read(hs.IdImpianto)
        If Not _i Is Nothing Then
            DesImpianto = _i.DesImpianto
            _i = Nothing
        End If
        mailContent = Replace(mailContent, "@!@DesImpianto@!@", DesImpianto)

        Dim Message As New System.Net.Mail.MailMessage
        Message.From = New MailAddress("noReply@clevergy.it")
        Message.To.Add("sistemi@clevergy.it")
        Message.To.Add("andrea.ravasio@clevergy.it")
        Message.To.Add("marco.fagnano@clevergy.it")
        Message.To.Add("andrea.laini@clevergy.it")
        'Message.To.Add("Lorenzo.Gritti@innowatio.it")
        'Message.To.Add("eduardo.giannarelli@innowatio.it")

        Dim _list As List(Of SCP.BLL.hs_UserAlert) = SCP.BLL.hs_UserAlert.List(hs.hsId)
        If Not _list Is Nothing Then
            For Each _obj As SCP.BLL.hs_UserAlert In _list
                If Not String.IsNullOrEmpty(_obj.emailaddress) Then
                    Message.To.Add(_obj.emailaddress)
                End If
            Next
            _list = Nothing
        End If

        Message.Subject = "Connessione ripristinata."
        Message.Body = mailContent
        Message.IsBodyHtml = True

        ' You won't need the calls to Message.Fields.Add()

        ' Replace SmtpMail.SmtpServer = server with the following:
        Dim smtp As New SmtpClient("mail.innowatio.it")

        'Dim smtp As New SmtpClient("smtp.mandrillapp.com")
        'smtp.Port = 587
        'smtp.EnableSsl = True
        'smtp.Credentials = New System.Net.NetworkCredential("clv.emcs.00@gmail.com", "gr4SmRnLxNn6hRFyW8lq-Q")
        Try
            smtp.Send(Message)

        Catch ex As Exception

        End Try
    End Sub

    Public Sub hsSendStatusAlarmON2clevergy(hs As SCP.BLL.HeatingSystem)
        Dim moduleName As String = "hsStatusAlarmON.html"
        Dim mailContent As String = loadModule(moduleName)
        mailContent = Replace(mailContent, "@!@hsId@!@", hs.hsId)
        mailContent = Replace(mailContent, "@!@Descr@!@", hs.Descr)
        mailContent = Replace(mailContent, "@!@lastRec@!@", FormatDateTime(hs.lastRec, DateFormat.GeneralDate))
        Dim DesImpianto As String = String.Empty
        Dim _i As SCP.BLL.Impianti = SCP.BLL.Impianti.Read(hs.IdImpianto)
        If Not _i Is Nothing Then
            DesImpianto = _i.DesImpianto
            _i = Nothing
        End If
        mailContent = Replace(mailContent, "@!@DesImpianto@!@", DesImpianto)

        Dim Message As New System.Net.Mail.MailMessage
        Message.From = New MailAddress("noReply@clevergy.it")
        Message.To.Add("sistemi@clevergy.it")
        Message.To.Add("andrea.ravasio@clevergy.it")
        Message.To.Add("marco.fagnano@clevergy.it")
        Message.To.Add("andrea.laini@clevergy.it")
        'Message.To.Add("Lorenzo.Gritti@innowatio.it")
        Dim _list As List(Of SCP.BLL.hs_UserAlert) = SCP.BLL.hs_UserAlert.List(hs.hsId)
        If Not _list Is Nothing Then
            For Each _obj As SCP.BLL.hs_UserAlert In _list
                If Not String.IsNullOrEmpty(_obj.emailaddress) Then
                    Message.To.Add(_obj.emailaddress)
                End If
            Next
            _list = Nothing
        End If

        Message.Subject = "(ALLARME) Sistema."
        Message.Body = mailContent
        Message.IsBodyHtml = True

        ' You won't need the calls to Message.Fields.Add()

        ' Replace SmtpMail.SmtpServer = server with the following:
        Dim smtp As New SmtpClient("mail.innowatio.it")

        'Dim smtp As New SmtpClient("smtp.mandrillapp.com")
        'smtp.Port = 587
        'smtp.EnableSsl = True
        'smtp.Credentials = New System.Net.NetworkCredential("clv.emcs.00@gmail.com", "gr4SmRnLxNn6hRFyW8lq-Q")
        Try
            smtp.Send(Message)

        Catch ex As Exception

        End Try
    End Sub

    Public Sub hsSendStatusAlarmOFF2clevergy(hs As SCP.BLL.HeatingSystem)
        Dim moduleName As String = "hsStatusAlarmOFF.html"
        Dim mailContent As String = loadModule(moduleName)
        mailContent = Replace(mailContent, "@!@hsId@!@", hs.hsId)
        mailContent = Replace(mailContent, "@!@Descr@!@", hs.Descr)
        mailContent = Replace(mailContent, "@!@lastRec@!@", FormatDateTime(hs.lastRec, DateFormat.GeneralDate))
        Dim DesImpianto As String = String.Empty
        Dim _i As SCP.BLL.Impianti = SCP.BLL.Impianti.Read(hs.IdImpianto)
        If Not _i Is Nothing Then
            DesImpianto = _i.DesImpianto
            _i = Nothing
        End If
        mailContent = Replace(mailContent, "@!@DesImpianto@!@", DesImpianto)

        Dim Message As New System.Net.Mail.MailMessage
        Message.From = New MailAddress("noReply@clevergy.it")
        Message.To.Add("sistemi@clevergy.it")
        Message.To.Add("andrea.ravasio@clevergy.it")
        Message.To.Add("marco.fagnano@clevergy.it")
        Message.To.Add("andrea.laini@clevergy.it")

        Dim _list As List(Of SCP.BLL.hs_UserAlert) = SCP.BLL.hs_UserAlert.List(hs.hsId)
        If Not _list Is Nothing Then
            For Each _obj As SCP.BLL.hs_UserAlert In _list
                If Not String.IsNullOrEmpty(_obj.emailaddress) Then
                    Message.To.Add(_obj.emailaddress)
                End If
            Next
            _list = Nothing
        End If

        Message.Subject = "(ALLARME RIENTRATO) Sistema."
        Message.Body = mailContent
        Message.IsBodyHtml = True

        ' You won't need the calls to Message.Fields.Add()

        ' Replace SmtpMail.SmtpServer = server with the following:
        Dim smtp As New SmtpClient("mail.innowatio.it")

        'Dim smtp As New SmtpClient("smtp.mandrillapp.com")
        'smtp.Port = 587
        'smtp.EnableSsl = True
        'smtp.Credentials = New System.Net.NetworkCredential("clv.emcs.00@gmail.com", "gr4SmRnLxNn6hRFyW8lq-Q")
        Try
            smtp.Send(Message)

        Catch ex As Exception

        End Try
    End Sub


    Public Function loadModule(moduleName As String) As String
        Dim fullpath As String = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "Modules\") & moduleName

        'Dim fullpath As String = HttpContext.Current.Request.MapPath("~\Modules\" & moduleName)
        Dim sr As StreamReader = New StreamReader(fullpath)
        Dim retVal As String = sr.ReadToEnd()
        sr.Close()
        Return retVal
    End Function
End Class
