'Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Outlook
'Imports Microsoft.Office.Interop

Public Class frmEmail

    Private g_EmailNames As Collection
    Private g_Emails As Collection
    Private g_EmailByName As Collection

    Private m_AutoCompleteStart As String = ""
    Private acscAutoComplete As AutoCompleteStringCollection

    Private m_oEmail As cEmail = Nothing

    'Private Sub frmEmail1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    '    Dim iCompte As Integer = 0
    '    Dim iPosition As Integer = 0
    '    Dim Address As AddressEntry = Nothing
    '    Dim oExchangeUser As ExchangeUser = Nothing
    '    Dim iErrorCount As Integer = 0

    '    Dim db As New cDBA
    '    Dim dt As DataTable

    '    If msOutlook Is Nothing Then
    '        msOutlook = CreateObject("Outlook.Application")
    '    End If

    '    'Dim Email As Microsoft.Office.Interop.Outlook.MailItem

    '    'Debug.Print("START LOADING")

    '    Try

    '        Dim strsql As String = ""

    '        Dim intNamePos As Integer

    '        iPosition = 0

    '        If g_Email Is Nothing Then
    '            g_Email = msOutlook.CreateItem(OlItemType.olMailItem)
    '            'm_Attachments = m_Email.Attachments
    '        End If

    '        'If g_Addresses Is Nothing Then
    '        '    iPosition = 1
    '        '    olSession = g_Email.Application.Session

    '        '    iPosition = 2
    '        '    g_Addresses = g_Email.Application.Session.GetGlobalAddressList()

    '        '    iPosition = 3
    '        '    g_EmailNames = New Collection
    '        '    g_Emails = New Collection
    '        '    g_EmailByName = New Collection

    '        '    iPosition = 4
    '        '    For Each Address In g_Addresses.AddressEntries
    '        '        'For Each Address As Object In g_Addresses.AddressEntries

    '        '        Dim strAddress As String = Address.Address.ToString
    '        '        intNamePos = InStrRev(strAddress.ToUpper, "CN=") + 3

    '        '        iPosition = 5
    '        '        'intNamePos = InStrRev(Address.Address.ToUpper, "CN=") + 3

    '        '        iPosition = 6
    '        '        'g_EmailNames.Add(Address.Name, Address.Name.ToUpper)

    '        '        iPosition = 7
    '        '        'g_Emails.Add(Address.Name, Trim(Mid(Address.Address.ToString, intNamePos).ToUpper))

    '        iPosition = 8
    '        strsql = _
    '        "SELECT ISNULL(Usr_ID, '') AS Usr_ID, ISNULL(FullName, '') AS FullName, ISNULL(Mail, '') AS Mail " & _
    '        "FROM   HUMRES WITH (Nolock) " & _
    '        "WHERE  Email IS NOT NULL ORDER BY FullName"

    '        'strsql = _
    '        '"SELECT ISNULL(FullName, '') AS FullName, ISNULL(Mail, '') AS Mail " & _
    '        '"FROM   HUMRES WITH (Nolock) " & _
    '        '"WHERE  UPPER(Usr_ID) = '" & SqlCompliantString(Trim(Mid(Address.Address.ToString, intNamePos).ToUpper)) & "' AND Email IS NOT NULL"

    '        dt = db.DataTable(strsql)
    '        iPosition = 9
    '        If dt.Rows.Count <> 0 Then

    '            For Each dtRow As DataRow In dt.Rows
    '                iPosition = 10

    '                g_EmailNames.Add(Trim(dt.Rows(0).Item("FullName").ToString), Trim(dt.Rows(0).Item("FullName").ToString).ToUpper)
    '                g_Emails.Add(Trim(dt.Rows(0).Item("FullName").ToString), Trim(Mid(strAddress, intNamePos).ToUpper))
    '                g_EmailByName.Add(Trim(dt.Rows(0).Item("Email").ToString), Trim(dt.Rows(0).Item("FullName").ToString).ToUpper)

    '                'Else

    '                'If iErrorCount < 4 Then

    '                '    iPosition = 1
    '                '    oExchangeUser = Address.GetExchangeUser
    '                '    iPosition = 12

    '                '    Try
    '                '        If oExchangeUser Is Nothing Then
    '                '            g_EmailByName.Add(Trim(Mid(Address.Address.ToString, intNamePos).ToString) & "@spectorandco.ca", Address.Name.ToUpper)
    '                '        Else
    '                '            g_EmailByName.Add(oExchangeUser.PrimarySmtpAddress.ToString, Address.Name.ToUpper)
    '                '        End If
    '                '    Catch er As System.Exception
    '                '        iErrorCount += 1
    '                '        If g_EmailNames.Contains(Address.Name.ToUpper) Then g_EmailNames.Remove(Address.Name.ToUpper)
    '                '        If g_EmailNames.Contains(Trim(Mid(Address.Address.ToString, intNamePos).ToUpper)) Then g_Emails.Remove(Trim(Mid(Address.Address.ToString, intNamePos).ToUpper))
    '                '    End Try

    '                'End If
    '            Next

    '        End If


    '        iPosition = 16

    '        'End If

    '        'Debug.Print("END loop")
    '        iPosition = 21

    '        acscAutoComplete = New AutoCompleteStringCollection

    '        For Each oName As String In g_EmailNames

    '            acscAutoComplete.Add(oName.ToString)

    '        Next

    '        iPosition = 22

    '        SyncLock txtTo.AutoCompleteCustomSource.SyncRoot
    '            txtTo.AutoCompleteCustomSource = acscAutoComplete
    '            txtCC.AutoCompleteCustomSource = acscAutoComplete
    '        End SyncLock

    '        iPosition = 23
    '        acscAutoComplete = Nothing

    '    Catch er As System.Exception
    '        If Address Is Nothing Then
    '            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '        Else
    '            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & vbCrLf & "Compte:" & Trim(iCompte.ToString) & "  Position:" & Trim(iPosition.ToString) & "  Address.Name: " & Address.Name.ToUpper)
    '        End If

    '    End Try

    '    'Debug.Print("DONE LOADING")

    'End Sub

    Private Sub frmEmail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim db As New cDBA
        Dim dt As DataTable

        Try

            Dim strsql As String = ""

            g_EmailNames = New Collection
            g_Emails = New Collection
            g_EmailByName = New Collection

            strsql =
            "SELECT     ISNULL(Usr_ID, '') AS Usr_ID, ISNULL(FullName, '') AS FullName, ISNULL(Mail, '') AS Mail " &
            "FROM       HUMRES WITH (Nolock) " &
            "WHERE      HUMRES.Mail IS NOT NULL AND (UPPER(HUMRES.Mail) LIKE '%SPECTORANDCO%' OR " &
            "UPPER(HUMRES.Mail) Like '%BANKERS%') AND HUMRES.Usr_ID <> 'ADMINISTRATOR' AND HUMRES.CONT_END_DATE IS NULL AND HUMRES.EMP_STAT <> 'I'" &
            "UNION      " &
            "SELECT     ISNULL(FullName, '') AS Usr_ID, ISNULL(FullName, '') AS FullName, ISNULL(Mail, '') AS Mail " &
            "FROM       OEI_ADDRESS_BOOK WITH (Nolock) " &
            "WHERE      OEI_ADDRESS_BOOK.Mail IS NOT NULL AND (UPPER(OEI_ADDRESS_BOOK.Mail) LIKE '%SPECTORANDCO%' OR UPPER(OEI_ADDRESS_BOOK.Mail) LIKE '%BANKERS%') " &
            "ORDER BY   FullName"


            '++ID 04232024 removed   HUMRES.CONT_END_DATE IS NULL AND    
            strsql =
            "SELECT     ISNULL(Usr_ID, '') AS Usr_ID, ISNULL(FullName, '') AS FullName, ISNULL(Mail, '') AS Mail " &
            "FROM       HUMRES WITH (Nolock) " &
            "WHERE      HUMRES.Mail IS NOT NULL AND (UPPER(HUMRES.Mail) LIKE '%SPECTORANDCO%' OR " &
            "UPPER(HUMRES.Mail) Like '%BANKERS%') AND HUMRES.Usr_ID <> 'ADMINISTRATOR' AND  HUMRES.EMP_STAT <> 'I'" &
            "UNION      " &
            "SELECT     ISNULL(FullName, '') AS Usr_ID, ISNULL(FullName, '') AS FullName, ISNULL(Mail, '') AS Mail " &
            "FROM       OEI_ADDRESS_BOOK WITH (Nolock) " &
            "WHERE      OEI_ADDRESS_BOOK.Mail IS NOT NULL AND (UPPER(OEI_ADDRESS_BOOK.Mail) LIKE '%SPECTORANDCO%' OR UPPER(OEI_ADDRESS_BOOK.Mail) LIKE '%BANKERS%') " &
            "ORDER BY   FullName"

            dt = db.DataTable(strsql)
            If dt.Rows.Count <> 0 Then

                For Each dtRow As DataRow In dt.Rows
                    'Debug.Print(Trim(dtRow.Item("Mail").ToString).ToUpper)
                    If Not g_EmailNames.Contains(Trim(dtRow.Item("FullName").ToString).ToUpper) Then
                        g_EmailNames.Add(Trim(dtRow.Item("FullName").ToString), Trim(dtRow.Item("FullName").ToString).ToUpper)
                        g_Emails.Add(Trim(dtRow.Item("FullName").ToString), Trim(dtRow.Item("Mail").ToString).ToUpper)
                        g_EmailByName.Add(Trim(dtRow.Item("Mail").ToString), Trim(dtRow.Item("FullName").ToString).ToUpper)
                    End If

                Next

            End If

            acscAutoComplete = New AutoCompleteStringCollection

            For Each oName As String In g_EmailNames

                acscAutoComplete.Add(oName.ToString)

            Next

            SyncLock txtTo.AutoCompleteCustomSource.SyncRoot
                txtTo.AutoCompleteCustomSource = acscAutoComplete
                txtCC.AutoCompleteCustomSource = acscAutoComplete
            End SyncLock

            acscAutoComplete = Nothing

            If m_oEmail Is Nothing Then
                txtSubject.Text = "Order #ORDER# : "
            Else
                txtTo.Text = m_oEmail.Email_To_Name
                txtCC.Text = m_oEmail.Email_CC_Name
                txtSubject.Text = m_oEmail.Subject
                txtBody.Text = m_oEmail.Body
            End If


        Catch er As System.Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' This sub checks if every entry exists.
    Private Sub CheckEmail(ByRef txtElement As xTextBox)

        Try
            'Dim Email As Microsoft.Office.Interop.Outlook.MailItem
            Dim txtListe As String = ""
            Dim strName As String
            Dim intBreakPos As Integer
            Dim strNames As String
            Dim blnFound As Boolean
            Dim blnSomeNotFound As Boolean = False

            Dim strFinalNames As String = ""
            Dim strNotFound As String = ""

            strNames = txtElement.Text

            'm_Email = msOutlook.CreateItem(OlItemType.olMailItem)

            While strNames <> ""

                intBreakPos = InStr(strNames, ";")
                If intBreakPos = 0 Then intBreakPos = strNames.Length

                strName = Mid(strNames, 1, intBreakPos).Replace(";", "").ToUpper

                blnFound = False

                If g_EmailNames.Contains(strName) Then

                    strFinalNames &= g_EmailNames.Item(strName).ToString & "; "
                    blnFound = True

                End If

                If Not blnFound Then
                    If g_Emails.Contains(strName.ToUpper) Then

                        strFinalNames &= g_Emails.Item(strName).ToString & "; "
                        blnFound = True

                    End If
                End If

                If Not blnFound Then
                    blnSomeNotFound = True
                    strNotFound &= Mid(strNames, 1, intBreakPos).Replace(";", "") & "; "
                End If

                If intBreakPos < Trim(strNames).Length Then

                    strNames = Trim(Mid(strNames, intBreakPos + 1))

                Else
                    strNames = ""
                End If

            End While

            If blnSomeNotFound Then

                MsgBox(strNotFound & " do not exist as names for emails.")

            End If

            txtElement.Text = strFinalNames

            'Email = Nothing

        Catch er As System.Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTo.Leave

        Call CheckEmail(txtTo)

    End Sub

    Private Sub txtCC_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCC.Leave

        Call CheckEmail(txtCC)

    End Sub

    ' Drag and drop has been disabled for txtMessage.
    'Private Sub txtMessage_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtBody.DragDrop

    '    Try

    '        Dim MyFiles() As String

    '        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

    '            ' Assign the files to an array.
    '            MyFiles = e.Data.GetData(DataFormats.FileDrop)

    '            AddDocumentFiles(MyFiles)

    '        ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then

    '            Dim objMI As Microsoft.Office.Interop.Outlook.MailItem
    '            ReDim MyFiles(1)

    '            MyFiles(0) = ""
    '            Dim iPos As Integer = 0
    '            For Each objMI In msOutlook.ActiveExplorer.Selection()
    '                ReDim Preserve MyFiles(iPos)
    '                'hardcode a destination path for testing
    '                Dim strFile As String = _
    '                            IO.Path.Combine("c:\ExactTemp", _
    '                                            (objMI.Subject + ".msg").Replace(":", "").Replace("/", "_").Replace("*", ".").Replace("?", "."))
    '                objMI.SaveAs(strFile)
    '                MyFiles(iPos) = strFile
    '                iPos += 1

    '            Next

    '            AddDocumentFiles(MyFiles)

    '        End If

    '    Catch er As System.Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    ' Drag and drop had been disabled for txtMessage
    'Private Sub txtMessage_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtBody.DragEnter

    '    If e.Data.GetDataPresent(DataFormats.FileDrop) Then
    '        e.Effect = DragDropEffects.All
    '    ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then 'For email Drag & Drop 
    '        e.Effect = DragDropEffects.Copy
    '    End If

    'End Sub

    ' Drag and drop had been disabled for txtMessage
    'Private Sub txtMessage_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtBody.DragOver

    '    e.Effect = DragDropEffects.Copy

    'End Sub

    'Private Sub AddDocumentFiles(ByRef Files As String())

    '    Dim lPos As Integer = 0
    '    Dim strFileName As String = ""

    '    Try
    '        'Dim oAttachments As Outlook.Attachments = Nothing

    '        'Loop through the array and add the files to the list.
    '        For Each strFile As String In Files

    '            If InStr(strFile, "\") <> 0 Then
    '                lPos = InStrRev(strFile, "\") + 1
    '            Else
    '                lPos = 1
    '            End If

    '            strFileName = Mid(strFile, lPos)
    '            m_Attachments.Add(strFile) ' , , m_Email.Body.Length + 1, strFileName)

    '            txtAttachments.Text += strFileName & "; "

    '        Next

    '    Catch er As System.Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    Finally
    '        'panProgress.Visible = False
    '    End Try

    'End Sub

    'Private Sub bgWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgWorker.DoWork

    '    Try
    '        Dim iCount As Integer = 0
    '        Dim h As IntPtr

    '        While iCount < 20 ' 1 second
    '            System.Threading.Thread.Sleep(50)

    '            h = FindWindow(Nothing, "Select Names: Global Address List")
    '            If h <> 0 Then
    '                Dim p() As Process = Process.GetProcessesByName("OUTLOOK") ' The Name of the Process You want to Minimize

    '                SetForegroundWindow(h)
    '                iCount = 21
    '            End If

    '            iCount += 1

    '        End While

    '    Catch er As System.Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Try
            Dim strName As String
            Dim strNames As String
            Dim strAddresses As String

            Dim intbreakpos As Integer = 0
            If Trim(txtTo.Text) = "" And Trim(txtCC.Text) = "" Then
                MsgBox("There must be at least one name or distibution list in the To, or CC Box.")
                Exit Sub
            End If

            If Trim(txtSubject.Text) = "" And Trim(txtBody.Text) = "" Then
                MsgBox("There must be a subject and a message to send.")
                Exit Sub
            End If


            Dim oMail As cEmail
            If m_oEmail Is Nothing Then
                oMail = New cEmail
            Else
                oMail = m_oEmail
            End If

            oMail.Email_From = Trim(g_User.Mail)

            strNames = txtTo.Text
            strAddresses = ""

            oMail.Email_To_Name = strNames

            While strNames <> ""

                intbreakpos = InStr(strNames, ";")
                If intbreakpos = 0 Then intbreakpos = strNames.Length

                strName = Mid(strNames, 1, intbreakpos).Replace(";", "").ToUpper

                If g_EmailByName.Contains(strName) Then

                    If strAddresses <> "" Then strAddresses &= "; "
                    strAddresses &= g_EmailByName.Item(strName).ToString ' & "@spectorandco.ca; "

                    'If oMail.Email_To_Name <> "" Then oMail.Email_To_Name &= "; "
                    'oMail.Email_To_Name &= Mid(strNames, 1, intbreakpos) ' .Replace(";", "")

                End If

                If intbreakpos < Trim(strNames).Length Then
                    strNames = Trim(Mid(strNames, intbreakpos + 1))
                Else
                    strNames = ""
                End If

            End While

            oMail.Email_To = strAddresses

            strNames = txtCC.Text
            strAddresses = ""

            oMail.Email_CC_Name = strNames

            While strNames <> ""

                intbreakpos = InStr(strNames, ";")
                If intbreakpos = 0 Then intbreakpos = strNames.Length

                strName = Mid(strNames, 1, intbreakpos).Replace(";", "").ToUpper

                If g_EmailByName.Contains(strName) Then

                    If strAddresses <> "" Then strAddresses &= "; "
                    strAddresses &= g_EmailByName.Item(strName).ToString ' & "@spectorandco.ca; "

                End If

                If intbreakpos < Trim(strNames).Length Then
                    strNames = Trim(Mid(strNames, intbreakpos + 1))
                Else
                    strNames = ""
                End If

            End While

            oMail.Email_CC = strAddresses

            If Trim(oMail.Email_To) = "" And Trim(oMail.Email_CC) = "" Then
                MsgBox("There must be at least one valid name or distibution list in the To, or CC Box.")
                Exit Sub
            End If

            oMail.Subject = txtSubject.Text
            oMail.Body = txtBody.Text

            oMail.UserID = g_User.Usr_ID
            oMail.CreateTS = Now
            oMail.Save()

            Me.Close()

        Catch er As System.Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Private Sub cmdCC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCC.Click

    '    Dim ms1Outlook As New Microsoft.Office.Interop.Outlook.Application

    '    Dim Email As Microsoft.Office.Interop.Outlook.MailItem
    '    Dim snd As SelectNamesDialog
    '    Dim txtListe As String = ""

    '    Email = ms1Outlook.CreateItem(OlItemType.olMailItem)
    '    snd = Email.Application.Session.GetSelectNamesDialog()

    '    snd.SetDefaultDisplayMode(OlDefaultSelectNamesDisplayMode.olDefaultMail)

    '    snd.Recipients.ResolveAll()

    '    bgWorker.RunWorkerAsync()

    '    snd.Display()

    '    SetForegroundWindow(Me.Handle)

    '    ' Add Recipients to meeting request.
    '    Dim recips As Microsoft.Office.Interop.Outlook.Recipients = snd.Recipients
    '    If recips.Count > 0 Then

    '        For Each recip As Microsoft.Office.Interop.Outlook.Recipient In recips
    '            'Debug.Print(recip.ToString)
    '            If recip.Type = 1 Then
    '                If InStr(txtTo.Text, recip.Name & "; ") = 0 Then
    '                    txtTo.Text &= recip.Name & "; "
    '                End If
    '            Else
    '                If InStr(txtCC.Text, recip.Name & "; ") = 0 Then
    '                    txtCC.Text &= recip.Name & "; "
    '                End If
    '            End If

    '            'appt.Recipients.Add(recip.Name)
    '            ' recip Type: 1 = To, 2 = CC, 3 = BCC
    '        Next

    '    End If

    '    Email = Nothing

    'End Sub

    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click
        'OEI-MAIL-MESS

        Try
            Dim oSearch As New cSearch("OEI-MAIL-MESS")

            If oSearch.Form.FoundRow <> -1 Then

                Dim dt As New DataTable
                Dim db As New cDBA

                Dim strSql As String = _
                "SELECT * " & _
                "FROM   OEI_Config_Messages WITH (Nolock) " & _
                "WHERE  MessageName = '" & oSearch.Form.FoundElementValue & "' "

                dt = db.DataTable(strSql)
                If dt.Rows.Count <> 0 Then

                    txtSubject.Text = dt.Rows(0).Item("Subject").ToString

                    'If Trim(txtBody.Text) = "" Then
                    txtBody.Text = dt.Rows(0).Item("Body").ToString
                    'Else
                    '    txtBody.Text &= " - - - - - - - - - - - - - - - - " & Chr(13) & Chr(13) & dt.Rows(0).Item("Body").ToString
                    'End If

                End If

            End If

            oSearch.Dispose()
            oSearch = Nothing

        Catch er As System.Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtSubject_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSubject.KeyDown

        Select Case e.KeyValue

            Case Keys.F7
                cmdSearch.PerformClick()

        End Select

    End Sub

    Public Sub SetEmailToAddress(ByVal pstrEmail As String)

        txtTo.Text = pstrEmail

    End Sub

    Public Sub SetEmailCCAddress(ByVal pstrEmail As String)

        txtCC.Text = pstrEmail

    End Sub

    Private Sub cmdTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTo.Click

        Try
            Dim oSearch As New cSearch("OEI-MAIL-USERS")

            If Trim(txtTo.Text).Length >= 1 Then
                If Mid(Trim(txtTo.Text), Trim(txtTo.Text).Length, 1) <> ";" Then txtTo.Text += "; "
            End If
            If Trim(oSearch.Form.FoundElementValue) <> "" Then txtTo.Text &= Trim(oSearch.Form.FoundElementValue) & "; "

        Catch er As System.Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdCC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCC.Click

        Try
            Dim oSearch As New cSearch("OEI-MAIL-USERS")

            If Trim(oSearch.Form.FoundElementValue) <> "" Then txtCC.Text &= Trim(oSearch.Form.FoundElementValue) & "; "

        Catch er As System.Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Property Email() As cEmail
        Get
            Email = m_oEmail
        End Get
        Set(ByVal value As cEmail)
            m_oEmail = value
        End Set
    End Property

End Class