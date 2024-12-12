Module mEmailer
    Public Function CreateEmailConnection() As Object
        Try
            'mailer = New MailBee.SMTP
            Dim mailer = CreateObject("MailBee.SMTP")

            ' Unlock MailBee.SMTP object
            mailer.LicenseKey = "MBC500-5959843B6F-3D8F3ABAB139E365D283A4325B857897"

            ' SMTP server name
            mailer.ServerName = "uranium"

            'real credentials (change to query later)
            mailer.ServerName = "exchange.bankers.ca"
            mailer.PortNumber = 366

            'Mailer.ServerName = "172.16.0.66"

            ' Mark that body has HTML format
            mailer.Message.BodyFormat = 1

            ' Enable logging SMTP session into a file. If any errors  
            ' occur, the log can be used to investigate the problem.
            mailer.EnableLogging = True
            mailer.LogFilePath = "C:\HV_Control_SendLog.txt"

            Return mailer
        Catch er As Exception
            Return vbNull
        End Try
    End Function

    Public Sub SendGenericEmail(ByVal from As String, ByVal addresses() As String, ByVal subject As String, ByVal body As String)
        Try

            'open mailer connection
            Dim mailer As Object = CreateEmailConnection()

            mailer.Message.FromAddr = "OEI Emailer Service <it@spectorandco.ca>"
            mailer.Message.ToAddr = addresses(0).ToString.Trim()
            If Not addresses(1) Is Nothing Then
                mailer.Message.CCAddr = addresses(1).ToString.Trim()
            End If

            If Not addresses(2) Is Nothing Then
                mailer.Message.BCCAddr = addresses(2).ToString.Trim()
            End If

            mailer.Message.Subject = subject.ToString.Trim()
            mailer.Message.BodyText = body.ToString.Trim()

            mailer.Send()

            'close connection
            mailer = Nothing

        Catch er As Exception
            MsgBox("Error in mEmailer->sendGenericEmail: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub
    Public Sub SendGenericEmail65(ByVal from As String, ByVal addresses() As String, ByVal subject As String, ByVal body As String)
        Try

            'open mailer connection
            Dim mailer As Object = CreateEmailConnection()

            mailer.Message.FromAddr = from '"OEI Emailer Service <it@spectorandco.ca>" '& "<it@spectorandco.ca>" '"OEI Emailer Service <it@spectorandco.ca>"
            mailer.Message.ToAddr = addresses(0).ToString.Trim()
            If Not addresses(1) Is Nothing Then
                mailer.Message.CCAddr = addresses(1).ToString.Trim()
            End If

            If Not addresses(2) Is Nothing Then
                mailer.Message.BCCAddr = addresses(2).ToString.Trim()
            End If

            mailer.Message.Subject = subject.ToString.Trim()
            mailer.Message.BodyText = body.ToString.Trim()

            If mailer.Send() Then
                MsgBox("Mail sent succesfully!!!")
            Else
                MsgBox("Mail not sent!!!")
            End If

            'close connection
            mailer = Nothing

        Catch er As Exception
            MsgBox("Error in mEmailer->sendGenericEmail: " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

End Module
