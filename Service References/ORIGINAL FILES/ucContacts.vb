Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports System.IO

Public Class ucContacts

    Private m_OrderGuid As String

    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////       VARIABLE DECLARATIONS             ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    ' Selected Contacts
    Private priSelIndex As Integer
    Private secSelIndex As Integer

    Private priClicked As Boolean
    Private secClicked As Boolean

    Private selectedIDInList As Integer
    Private selectedMethodInList As String
    Private primarySelected As Short

    Private dbConn As cDBA
    Private dtContacts As DataTable

    'Private rsCustomerInfo As New ADODB.Recordset

    'Private m_oOrder As New cOrder
    Private blnLoading As Boolean = False

    ' This flag is set to catch the _optType_0.CheckedChanged that occurs when loading fromOrder.
    ' We must check why that happens and fix, to remove the flag.
    Private blnInit As Boolean = True

    Private m_Contact As cOrderContacts

    Private Enum Columns
        FullName
        Fax
        Email
        ID
    End Enum

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        m_Contact = New cOrderContacts

        blnInit = False

    End Sub

    Public Sub Fill() ' ByRef pOrder As cOrder)

        blnLoading = True
        'm_oOrder = pOrder

        Try

            If Trim(m_oOrder.Ordhead.Ord_GUID).Length <> 0 Then

                txtCusNo.Text = m_oOrder.Ordhead.Customer.Cmp_Code
                txtCusName.Text = m_oOrder.Ordhead.Customer.Cmp_Name

                With m_oOrder.Ordhead.Contacts

                    txtOEContact.TextDataID = .OEContact.ID
                    txtOEContact.Text = .OEContact.FullName '  .ContactName(.OEContact.ID)

                    If m_oOrder.Ordhead.User_Def_Fld_4.Trim = String.Empty Then m_oOrder.Ordhead.User_Def_Fld_4 = .OEContact.FullName

                    txtPaperproof.TextDataID = .Paperproof.ID
                    txtPaperproof.Text = .Paperproof.FullName ' .ContactName(.Paperproof.ID)

                    txtOrderAck.TextDataID = .OrderAck.ID
                    txtOrderAck.Text = .OrderAck.FullName ' .ContactName(.OrderAck.ID)

                    txtShipping.TextDataID = .Shipping.ID
                    txtShipping.Text = .Shipping.FullName ' .ContactName(.Shipping.ID)

                    txtPreShipping.TextDataID = .PreShipping.ID
                    txtPreShipping.Text = .PreShipping.FullName ' .ContactName(.PreShipping.ID)

                    txt90DayRemind.TextDataID = .Reminder90Day.ID
                    txt90DayRemind.Text = .Reminder90Day.FullName ' .ContactName(.Reminder90Day.ID)

                End With

            Else
                Reset() ' pOrder)
            End If

            Call LoadNewContactComboBoxes()
            Call LoadGrdContacts()

            txtOEContact.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        blnLoading = False

    End Sub

    Public Sub Reset() ' ByRef pOrder As cOrder)

        txtOEContact.TextInit()
        txtPaperproof.TextInit()
        txtOrderAck.TextInit()
        txtShipping.TextInit()
        txtPreShipping.TextInit()
        txt90DayRemind.TextInit()

    End Sub

    Public Sub Save()

        Try

            If m_oOrder.Ordhead.Contacts Is Nothing Then Exit Sub

            m_oOrder.Ordhead.Contacts.Save()

            'With m_oOrder.Ordhead.Contacts

            '    .OEContact.ID = txtOEContact.TextDataID
            '    .Paperproof.ID = txtPaperproof.TextDataID
            '    .OrderAck.ID = txtOrderAck.TextDataID
            '    .Shipping.ID = txtShipping.TextDataID
            '    .PreShipping.ID = txtPreShipping.TextDataID
            '    .Reminder90Day.ID = txt90DayRemind.TextDataID

            'End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public ReadOnly Property Order() As cOrder
        Get
            Order = m_oOrder
        End Get
    End Property

    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////       GENERAL PROCEDURES            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub LoadGrdContacts()

        Try
            If txtCusNo.Text = "" Then Exit Sub

            Dim strCmp_WWN As String = ""
            Dim dtCmpy As New DataTable
            dbConn = New cDBA()

            Dim strSql As String = _
            " SELECT    * " & _
            " FROM      cicmpy cmp WITH (Nolock) " & _
            " WHERE     cmp.cmp_code = '" & txtCusNo.Text & "' AND " & _
            "           cmp.cmp_type = 'C' "

            dtCmpy = dbConn.DataTable(strSql)

            If dtCmpy.Rows.Count <> 0 Then
                strCmp_WWN = dtCmpy.Rows(0).Item("Cmp_WWN").ToString
            End If
            dtCmpy.Dispose()

            strSql = _
            " SELECT    FullName, RTRIM(cnt_f_fax) AS cnt_f_fax, RTRIM(cnt_email) AS cnt_email, ID " & _
            " FROM      cicntp WITH (Nolock) " & _
            " WHERE     cmp_wwn =  '" & strCmp_WWN & "' " & _
            " ORDER BY  FullName "

            '"SELECT ID, Cmt AS Comments, Cus_Sort " & _
            '"FROM   Exact_Custom_OeFlash_CustomerComments WITH (Nolock) " & _
            '"WHERE  Cus_no = '" & txtCusNo.Text & "' ORDER BY Cus_Sort "

            'daDataAdapter = New OleDb.OleDbDataAdapter
            dtContacts = New DataTable

            dbConn = New cDBA()

            'daDataAdapter = dbConn.DataAdapter(sql)
            dtContacts = dbConn.DataTable(strSql)
            dgvContacts.DataSource = dtContacts

            Call SetGrdContactsColumns()

            If gSelectedType > 0 Then
                optType(gSelectedType).Checked = False
                optType(gSelectedType).Checked = True
            Else
                optType(ContactType.MainContact).Checked = False
                optType(ContactType.MainContact).Checked = True
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SetGrdContactsColumns()

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

            'With grdContacts
            With dgvContacts

                With .Columns(0)
                    .HeaderText = "Name"
                    .Width = 150
                    '.DefaultCellStyle = MSDataGridLib.AlignmentConstants.dbgLeft
                End With
                With .Columns(1)
                    .HeaderText = "Fax"
                    .Width = 105
                    '.Alignment = MSDataGridLib.AlignmentConstants.dbgLeft
                End With
                With .Columns(2)
                    .HeaderText = "Email"
                    .Width = 300
                    '.Alignment = MSDataGridLib.AlignmentConstants.dbgLeft
                End With
                With .Columns(3)
                    .HeaderText = "ID"
                    .Width = 0
                    .Visible = False
                End With

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////       SAVE / GET FORM POSITION             ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub dgvContacts_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvContacts.CellBeginEdit

        ' We don't want to overwrite the ID field or the contact name
        If e.ColumnIndex = Columns.FullName Or e.ColumnIndex = Columns.ID Then e.Cancel = True

    End Sub
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#Region "Add contacts section #############################################"

    Private Sub LoadNewContactComboBoxes()

        Dim db As New cDBA
        Dim dtNamePreds As DataTable
        Dim dtLangs As DataTable

        'Dim rsNamePreds As New ADODB.Recordset
        'Dim rsLangs As New ADODB.Recordset
        Dim strSql As String

        Try

            ' Load Name Preds
            strSql = _
            "SELECT     PredCode " & _
            "FROM       Pred WITH (nolock) " & _
            "WHERE      PredCode IS NOT NULL " & _
            "ORDER BY   ID "

            dtNamePreds = db.DataTable(strSql)
            'rsNamePreds.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            For Each dtRow As DataRow In dtNamePreds.Rows
                cboPred.Items.Add(dtRow.Item("PredCode").ToString().Trim())
            Next

            'If rsNamePreds.RecordCount <> 0 Then
            '    While Not rsNamePreds.EOF
            '        cboPred.Items.Add(rsNamePreds.Fields("PredCode").Value)
            '        rsNamePreds.MoveNext()
            '    End While
            'End If
            'rsNamePreds.Close()

            cboPred.Text = "MR"

            ' Load Languages
            strSql = _
            "       SELECT 		1 AS TaalOrder, 'FR' AS TaalCode " & _
            "UNION  SELECT 		1 AS TaalOrder, 'EN' AS TaalCode " & _
            "UNION  SELECT     	2 as TaalOrder, TaalCode  " & _
            "       FROM       	Taal WITH (Nolock) " & _
            "       WHERE       TaalCode IS NOT NULL " & _
            "       ORDER BY   	TaalOrder, TaalCode "

            dtLangs = db.DataTable(strSql)
            If dtLangs.Rows.Count <> 0 Then
                For Each dtRow As DataRow In dtLangs.Rows
                    cboLang.Items.Add(dtRow.Item("TaalCode").ToString().Trim())
                Next
            End If
            'rsLangs.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            'If rsLangs.RecordCount <> 0 Then
            '    While Not rsLangs.EOF
            '        cboLang.Items.Add(rsLangs.Fields("TaalCode").Value)
            '        rsLangs.MoveNext()
            '    End While
            'End If
            'rsLangs.Close()
            cboLang.Text = "EN"

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub cmdInsertNewContact_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInsertNewContact.Click

        Dim strSql As String = ""
        'Dim rsContactInfo As New ADODB.Recordset
        Dim db As New cDBA
        Dim dt As DataTable

        Try
            'Dim sqlShip, sql, sqlPics As Object
            Dim strFullName As String
            Dim gender As String

            If Trim(cboPred.Text) = "" Then Throw New OEException(OEError.Contact_Name_Pred_Needed)
            If Trim(txtFirstName.Text) = "" Then Throw New OEException(OEError.Contact_First_Name_Needed)
            If Trim(txtLastName.Text) = "" Then Throw New OEException(OEError.Contact_Last_Name_Needed)
            If Trim(cboLang.Text) = "" Then Throw New OEException(OEError.Contact_Preferred_Language_Needed)

            If Trim(txtTel.Text) <> "" And Not IsDBNull(Trim(txtTel.Text)) Then
                If Not m_oOrder.Validation.IsValidPhoneNumber(Trim(txtTel.Text)) Then Throw New OEException(OEError.Contact_Incorrect_Phone_Syntax)
            End If

            If Trim(txtFax.Text) <> "" And Not IsDBNull(Trim(txtFax.Text)) Then
                If Not m_oOrder.Validation.IsValidPhoneNumber(Trim(txtFax.Text)) Then Throw New OEException(OEError.Contact_Incorrect_Fax_Syntax)
            End If

            If Trim(txtEmail.Text) <> "" And Not IsDBNull(Trim(txtEmail.Text)) Then
                If Not m_oOrder.Validation.IsValidEmail(Trim(txtEmail.Text)) Then Throw New OEException(OEError.Contact_Incorrect_Email)
            End If

            If txtTel.Text = "" And txtFax.Text = "" And txtEmail.Text = "" Then
                Throw New OEException(OEError.Contact_Method_Needed)
            End If

            If UCase(cboPred.Text) = "MR" Or UCase(cboPred.Text) = "MSR" Then
                gender = "M"
            ElseIf UCase(cboPred.Text) = "MRS" Or UCase(cboPred.Text) = "MME" Or UCase(cboPred.Text) = "MS" Then
                gender = "F"
            Else
                gender = "M"
            End If

            ' Add Contact to 'cicntp' table
            strFullName = Replace(Trim(txtFirstName.Text), "'", "''")
            strFullName = strFullName & " " & Replace(Trim(txtLastName.Text), "'", "''")

            Dim oContact As New cContact(m_OrderGuid, ContactType.MainContact)

            With m_oOrder.Ordhead.Contacts

                If _optType_0.Checked Then
                    oContact = .OEContact
                ElseIf _optType_1.Checked Then
                    oContact = .Paperproof
                ElseIf _optType_2.Checked Then
                    oContact = .PreShipping
                ElseIf _optType_3.Checked Then
                    oContact = .Shipping
                ElseIf _optType_4.Checked Then
                    oContact = .OrderAck
                ElseIf _optType_5.Checked Then
                    oContact = .Reminder90Day
                End If

            End With

            oContact.Customer = m_oOrder.Ordhead.Customer
            oContact.FirstName = txtFirstName.Text
            oContact.LastName = txtLastName.Text
            oContact.FullName = strFullName
            oContact.Cnt_Email = txtEmail.Text
            oContact.Cnt_F_Tel = txtTel.Text
            oContact.Cnt_F_Fax = txtFax.Text
            oContact.Cnt_F_Ext = txtExt.Text
            oContact.PredCode = cboPred.Text
            oContact.TaalCode = cboLang.Text
            oContact.Cnt_Acc_Man = m_oOrder.Ordhead.Customer.Cmp_Acc_Man

            oContact.Save()

            'Dim rs As New ADODB.Recordset
            'Dim intContactId As Integer
            Dim strContactName As String = ""

            strSql = " SELECT * FROM cicntp WITH (Nolock) WHERE cmp_wwn = '" & m_oOrder.Ordhead.Customer.Cmp_WWN & "' ORDER BY ID DESC "
            'rs.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            dt = db.DataTable(strSql)

            'If rs.RecordCount <> 0 Then
            If dt.Rows.Count <> 0 Then
                'intContactId = rs.Fields("ID").Value
                'strContactName = oContact.FullName
                'oContact.ID = rs.Fields("ID").Value
                oContact.ID = dt.Rows(0).Item("ID")
            End If
            'rs.Close()

            If _optType_0.Checked And txtOEContact.TextDataID <> oContact.ID Then
                txtOEContact.Text = oContact.FullName
                txtOEContact.TextDataID = oContact.ID
            End If

            If _optType_1.Checked And txtPaperproof.TextDataID <> oContact.ID Then
                txtPaperproof.Text = oContact.FullName
                txtPaperproof.TextDataID = oContact.ID

            ElseIf _optType_2.Checked And txtPreShipping.TextDataID <> oContact.ID Then
                txtPreShipping.Text = oContact.FullName
                txtPreShipping.TextDataID = oContact.ID

            ElseIf _optType_3.Checked And txtShipping.TextDataID <> oContact.ID Then
                txtShipping.Text = oContact.FullName
                txtShipping.TextDataID = oContact.ID

            ElseIf _optType_4.Checked And txtOrderAck.TextDataID <> oContact.ID Then
                txtOrderAck.Text = oContact.FullName
                txtOrderAck.TextDataID = oContact.ID

            ElseIf _optType_5.Checked And txt90DayRemind.TextDataID <> oContact.ID Then
                txt90DayRemind.Text = oContact.FullName
                txt90DayRemind.TextDataID = oContact.ID

            End If

            MsgBox("New Contact Added.")

            Call LoadGrdContacts()

            cboLang.Text = "EN"
            cboPred.Text = "MR"
            txtFirstName.Text = ""
            txtLastName.Text = ""
            txtTel.Text = ""
            txtExt.Text = ""
            txtFax.Text = ""
            txtEmail.Text = ""

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            'Finally
            '    If rsContactInfo.State <> 0 Then rsContactInfo.Close()
        End Try

    End Sub

#End Region

#Region "Contact Textboxes GotFocus #######################################"

    Private Sub txtOEContact_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOEContact.GotFocus
        _optType_0.Checked = True
    End Sub

    Private Sub txtPaperproof_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaperproof.GotFocus
        _optType_1.Checked = True
    End Sub

    Private Sub txtOrderAck_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOrderAck.GotFocus
        _optType_4.Checked = True
    End Sub

    Private Sub txtShipping_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShipping.GotFocus
        _optType_3.Checked = True
    End Sub

    Private Sub txtPreShipping_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPreShipping.GotFocus
        _optType_2.Checked = True
    End Sub

    Private Sub txt90DayRemind_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt90DayRemind.GotFocus
        _optType_5.Checked = True
    End Sub

#End Region

#Region "Radio buttons Checked Changed events #############################"

    Private Sub _optType_0_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _optType_0.CheckedChanged
        If blnInit Then Exit Sub
        If _optType_0.Checked And Not m_oOrder.Ordhead.Contacts Is Nothing Then Call PreSelectContact(m_oOrder.Ordhead.Contacts.OEContact.ID) ' txtOEContact.TextDataID)
        'If _optType_0.Checked Then MsgBox("x") ' Call PreSelectContact(m_oOrder.Ordhead.Contacts.OEContact.ID) ' txtOEContact.TextDataID)
    End Sub
    Private Sub _optType_1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _optType_1.CheckedChanged
        If _optType_1.Checked Then Call PreSelectContact(m_oOrder.Ordhead.Contacts.Paperproof.ID) ' txtPaperproof.TextDataID)
    End Sub

    Private Sub _optType_4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _optType_4.CheckedChanged
        If _optType_4.Checked Then Call PreSelectContact(m_oOrder.Ordhead.Contacts.OrderAck.ID) ' txtOrderAck.TextDataID)
    End Sub

    Private Sub _optType_3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _optType_3.CheckedChanged
        If _optType_3.Checked Then Call PreSelectContact(m_oOrder.Ordhead.Contacts.Shipping.ID) ' txtShipping.TextDataID)
    End Sub

    Private Sub _optType_2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _optType_2.CheckedChanged
        If _optType_2.Checked Then Call PreSelectContact(m_oOrder.Ordhead.Contacts.PreShipping.ID) ' txtPreShipping.TextDataID)
    End Sub

    Private Sub _optType_5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _optType_5.CheckedChanged
        If _optType_5.Checked Then Call PreSelectContact(m_oOrder.Ordhead.Contacts.Reminder90Day.ID) ' txt90DayRemind.TextDataID)
    End Sub

    Private Sub PreSelectContact(ByVal pintContactID As Integer)

        Try
            If pintContactID = 0 Then
                optEmail.Checked = False
                optFax.Checked = False
                optPrimary.Checked = False
                optSecondary.Checked = False
                Exit Sub
            Else
                Dim intPos As Integer
                For intPos = 0 To dgvContacts.Rows.Count - 1
                    If dgvContacts.Rows(intPos).Cells(Columns.ID).Value = pintContactID Then
                        dgvContacts.CurrentCell = dgvContacts.Rows(intPos).Cells(Columns.FullName)
                        Exit For
                    End If
                Next
            End If

            Dim oContact As New cContact
            Dim intMethod As ContactMethod
            Dim intInfo As ContactInfo

            If _optType_0.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.OEContact

            ElseIf _optType_1.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.Paperproof

            ElseIf _optType_2.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.PreShipping

            ElseIf _optType_3.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.Shipping

            ElseIf _optType_4.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.OrderAck

            ElseIf _optType_5.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.Reminder90Day

            End If

            intMethod = oContact.OrderContact.DefMethod
            intInfo = oContact.OrderContact.PrimaryContact

            optEmail.Checked = (intMethod = ContactMethod.Email)
            optFax.Checked = (intMethod = ContactMethod.Fax)

            optPrimary.Checked = (intInfo = ContactInfo.Primary)
            optSecondary.Checked = (intInfo = ContactInfo.Secondary)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub PreSelectContact(ByVal pContact As cContact)

        Try

            If pContact.ID = 0 Then
                optEmail.Checked = False
                optFax.Checked = False
                optPrimary.Checked = False
                optSecondary.Checked = False
                Exit Sub
            Else
                Dim intPos As Integer
                For intPos = 0 To dgvContacts.Rows.Count - 1
                    If dgvContacts.Rows(intPos).Cells(Columns.ID).Value = pContact.ID Then
                        dgvContacts.CurrentCell = dgvContacts.Rows(intPos).Cells(Columns.FullName)
                        Exit For
                    End If
                Next
            End If

            'Dim oContact As New cContact
            Dim intMethod As ContactMethod
            Dim intInfo As ContactInfo

            'If _optType_0.Checked Then
            '    oContact = m_oOrder.Ordhead.Contacts.OEContact

            'ElseIf _optType_1.Checked Then
            '    oContact = m_oOrder.Ordhead.Contacts.Paperproof

            'ElseIf _optType_2.Checked Then
            '    oContact = m_oOrder.Ordhead.Contacts.PreShipping

            'ElseIf _optType_3.Checked Then
            '    oContact = m_oOrder.Ordhead.Contacts.Shipping

            'ElseIf _optType_4.Checked Then
            '    oContact = m_oOrder.Ordhead.Contacts.OrderAck

            'ElseIf _optType_5.Checked Then
            '    oContact = m_oOrder.Ordhead.Contacts.Reminder90Day

            'End If

            intMethod = pContact.OrderContact.DefMethod ' IIf(pContact.Order.DefMethod = "", 0, pContact.Order.DefMethod)
            intInfo = pContact.OrderContact.PrimaryContact

            optEmail.Checked = (intMethod = ContactMethod.Email)
            optFax.Checked = (intMethod = ContactMethod.Fax)

            optPrimary.Checked = (intInfo = ContactInfo.Primary)
            optSecondary.Checked = (intInfo = ContactInfo.Secondary)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


#End Region

    Private Sub dgvContacts_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvContacts.CellValidating

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        If dgvContacts.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString = e.FormattedValue.ToString Then Exit Sub

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String = "SELECT * FROM Cicntp WHERE ID = " & dgvContacts.Rows(e.RowIndex).Cells(Columns.ID).Value.ToString

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                drRow = dt.Rows(0)

                Select Case e.ColumnIndex
                    Case Columns.Email
                        drRow.Item("Cnt_Email") = e.FormattedValue

                    Case Columns.Fax
                        drRow.Item("Cnt_F_Fax") = e.FormattedValue
                End Select

                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand

            End If
            'drRow = IIf(dt.Rows.Count = 0, dt.NewRow(), dt.Rows(0))

            db.DBDataAdapter.Update(db.DBDataTable)

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdRemoveContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRemoveContact.Click

        Try

            Dim oTextbox As New xTextBox
            Dim oContact As New cContact()

            If _optType_0.Checked Then
                oTextbox = txtOEContact
                oContact = m_oOrder.Ordhead.Contacts.OEContact
            ElseIf _optType_1.Checked Then
                oTextbox = txtPaperproof
                oContact = m_oOrder.Ordhead.Contacts.Paperproof
            ElseIf _optType_2.Checked Then
                oTextbox = txtPreShipping
                oContact = m_oOrder.Ordhead.Contacts.PreShipping
            ElseIf _optType_3.Checked Then
                oTextbox = txtShipping
                oContact = m_oOrder.Ordhead.Contacts.Shipping
            ElseIf _optType_4.Checked Then
                oTextbox = txtOrderAck
                oContact = m_oOrder.Ordhead.Contacts.OrderAck
            ElseIf _optType_5.Checked Then
                oTextbox = txt90DayRemind
                oContact = m_oOrder.Ordhead.Contacts.Reminder90Day
            End If

            oTextbox.Text = ""
            oTextbox.TextDataID = 0
            oContact.ID = 0

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdCopyContacts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyContacts.Click

        Dim strContactName As String = ""
        Dim intContactID As Integer = 0
        Dim intMethod As ContactMethod
        Dim intInfo As ContactInfo

        Try
            Dim oContact As New cContact
            Dim oTextBox As New xTextBox

            With m_oOrder.Ordhead.Contacts
                If _optType_0.Checked Then
                    oContact = .OEContact

                ElseIf _optType_1.Checked Then
                    oContact = .Paperproof

                ElseIf _optType_2.Checked Then
                    oContact = .PreShipping

                ElseIf _optType_3.Checked Then
                    oContact = .Shipping

                ElseIf _optType_4.Checked Then
                    oContact = .OrderAck

                ElseIf _optType_5.Checked Then
                    oContact = .Reminder90Day

                End If

                intContactID = oContact.ID
                If intContactID = 0 Then
                    oContact.FullName = ""
                Else
                    strContactName = oContact.FullName
                End If
                intMethod = oContact.OrderContact.DefMethod
                intInfo = oContact.OrderContact.PrimaryContact

                For iPos As Integer = ContactType.MainContact To ContactType.Reminder90day

                    Select Case iPos
                        Case ContactType.MainContact
                            oContact = .OEContact
                            oTextBox = txtOEContact
                        Case ContactType.OrderAck
                            oContact = .OrderAck
                            oTextBox = txtOrderAck
                        Case ContactType.Paperproof
                            oContact = .Paperproof
                            oTextBox = txtPaperproof
                        Case ContactType.PreShipping
                            oContact = .PreShipping
                            oTextBox = txtPreShipping
                        Case ContactType.Reminder90day
                            oContact = .Reminder90Day
                            oTextBox = txt90DayRemind
                        Case ContactType.Shipping
                            oContact = .Shipping
                            oTextBox = txtShipping

                    End Select

                    oContact.FullName = strContactName
                    oContact.ID = intContactID
                    oContact.OrderContact.DefMethod = intMethod
                    oContact.OrderContact.PrimaryContact = intInfo

                    oTextBox.Text = oContact.FullName
                    oTextBox.TextDataID = oContact.ID

                Next

            End With

            'If selectedIDInList < 0 Then
            '    MsgBox("You must first select a contact to copy!")
            '    Exit Sub
            'End If

            'answer = MsgBox("Are you sure you want to contact this contact for all types?", 65, "Copy Contact")

            'If answer <> 1 Then
            '    Exit Sub
            'Else

            '    For i = ContactType.MainContact To ContactType.Reminder90day
            '        Call InsertOrUpdateUser(selectedIDInList, selectedMethodInList, i, primarySelected)
            '    Next i

            'End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub optPrimary_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPrimary.CheckedChanged

        If m_oOrder.Ordhead.Contacts Is Nothing Then Exit Sub

        With m_oOrder.Ordhead.Contacts

            If _optType_0.Checked Then
                .OEContact.OrderContact.PrimaryContact = ContactInfo.Primary

            ElseIf _optType_1.Checked Then
                .Paperproof.OrderContact.PrimaryContact = ContactInfo.Primary

            ElseIf _optType_2.Checked Then
                .PreShipping.OrderContact.PrimaryContact = ContactInfo.Primary

            ElseIf _optType_3.Checked Then
                .Shipping.OrderContact.PrimaryContact = ContactInfo.Primary

            ElseIf _optType_4.Checked Then
                .OrderAck.OrderContact.PrimaryContact = ContactInfo.Primary

            ElseIf _optType_5.Checked Then
                .Reminder90Day.OrderContact.PrimaryContact = ContactInfo.Primary

            End If

        End With

    End Sub

    Private Sub optSecondary_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSecondary.CheckedChanged

        Try
            If m_oOrder.Ordhead.Contacts Is Nothing Then Exit Sub

            With m_oOrder.Ordhead.Contacts

                If _optType_0.Checked Then
                    .OEContact.OrderContact.PrimaryContact = ContactInfo.Secondary

                ElseIf _optType_1.Checked Then
                    .Paperproof.OrderContact.PrimaryContact = ContactInfo.Secondary

                ElseIf _optType_2.Checked Then
                    .PreShipping.OrderContact.PrimaryContact = ContactInfo.Secondary

                ElseIf _optType_3.Checked Then
                    .Shipping.OrderContact.PrimaryContact = ContactInfo.Secondary

                ElseIf _optType_4.Checked Then
                    .OrderAck.OrderContact.PrimaryContact = ContactInfo.Secondary

                ElseIf _optType_5.Checked Then
                    .Reminder90Day.OrderContact.PrimaryContact = ContactInfo.Secondary

                End If

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub optFax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optFax.CheckedChanged

        Try
            If m_oOrder.Ordhead.Contacts Is Nothing Then Exit Sub

            Dim intMethode As Integer = ContactMethod.None

            If optEmail.Checked Then
                intMethode = ContactMethod.Email
            ElseIf optFax.Checked Then
                intMethode = ContactMethod.Fax
            Else
                intMethode = ContactMethod.None
            End If

            With m_oOrder.Ordhead.Contacts

                If _optType_0.Checked Then
                    .OEContact.OrderContact.DefMethod = intMethode

                ElseIf _optType_1.Checked Then
                    .Paperproof.OrderContact.DefMethod = intMethode

                ElseIf _optType_2.Checked Then
                    .PreShipping.OrderContact.DefMethod = intMethode

                ElseIf _optType_3.Checked Then
                    .Shipping.OrderContact.DefMethod = intMethode

                ElseIf _optType_4.Checked Then
                    .OrderAck.OrderContact.DefMethod = intMethode

                ElseIf _optType_5.Checked Then
                    .Reminder90Day.OrderContact.DefMethod = intMethode

                End If

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub optEmail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optEmail.CheckedChanged

        Try

            Dim intMethode As Integer = ContactMethod.None

            If optEmail.Checked Then
                intMethode = ContactMethod.Email
            ElseIf optFax.Checked Then
                intMethode = ContactMethod.Fax
            Else
                intMethode = ContactMethod.None
            End If

            With m_oOrder.Ordhead.Contacts

                If _optType_0.Checked Then
                    .OEContact.OrderContact.DefMethod = intMethode

                ElseIf _optType_1.Checked Then
                    .Paperproof.OrderContact.DefMethod = intMethode

                ElseIf _optType_2.Checked Then
                    .PreShipping.OrderContact.DefMethod = intMethode

                ElseIf _optType_3.Checked Then
                    .Shipping.OrderContact.DefMethod = intMethode

                ElseIf _optType_4.Checked Then
                    .OrderAck.OrderContact.DefMethod = intMethode

                ElseIf _optType_5.Checked Then
                    .Reminder90Day.OrderContact.DefMethod = intMethode

                End If

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvContacts.DoubleClick

        Dim intContactID As Integer = dgvContacts.CurrentRow.Cells(Columns.ID).Value
        Dim strContactName As String = dgvContacts.CurrentRow.Cells(Columns.FullName).Value
        Dim oCicntp As New cCicntp(intContactID)

        If dgvContacts.Rows.Count = 0 Then Exit Sub

        Try
            Dim oContact As New cContact
            Dim oTextBox As New xTextBox
            Dim oNextTextBox As New xTextBox

            If _optType_0.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.OEContact
                oTextBox = txtOEContact
                oNextTextBox = txtPaperproof

                If m_oOrder.Ordhead.Contacts.OEContact.ID = 0 Then

                    txtPaperproof.Text = strContactName
                    txtPaperproof.TextDataID = intContactID
                    m_oOrder.Ordhead.Contacts.Paperproof.ID = intContactID
                    m_oOrder.Ordhead.Contacts.Paperproof.FullName = strContactName

                    txtPreShipping.Text = strContactName
                    txtPreShipping.TextDataID = intContactID
                    m_oOrder.Ordhead.Contacts.PreShipping.ID = intContactID
                    m_oOrder.Ordhead.Contacts.PreShipping.FullName = strContactName

                    txtShipping.Text = strContactName
                    txtShipping.TextDataID = intContactID
                    m_oOrder.Ordhead.Contacts.Shipping.ID = intContactID
                    m_oOrder.Ordhead.Contacts.Shipping.FullName = strContactName

                    txtOrderAck.Text = strContactName
                    txtOrderAck.TextDataID = intContactID
                    m_oOrder.Ordhead.Contacts.OrderAck.ID = intContactID
                    m_oOrder.Ordhead.Contacts.OrderAck.FullName = strContactName

                    txt90DayRemind.Text = strContactName
                    txt90DayRemind.TextDataID = intContactID
                    m_oOrder.Ordhead.Contacts.Reminder90Day.ID = intContactID
                    m_oOrder.Ordhead.Contacts.Reminder90Day.FullName = strContactName

                    If oCicntp.Use_Account_No.Trim <> String.Empty Then
                        MsgBox("Customer requires that this contact use account # " & oCicntp.Use_Account_No.Trim & ". ")
                    End If

                End If

            ElseIf _optType_1.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.Paperproof
                oTextBox = txtPaperproof
                oNextTextBox = txtOrderAck

            ElseIf _optType_2.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.PreShipping
                oTextBox = txtPreShipping
                oNextTextBox = txt90DayRemind

            ElseIf _optType_3.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.Shipping
                oTextBox = txtShipping
                oNextTextBox = txtPreShipping

            ElseIf _optType_4.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.OrderAck
                oTextBox = txtOrderAck
                oNextTextBox = txtShipping

            ElseIf _optType_5.Checked Then
                oContact = m_oOrder.Ordhead.Contacts.Reminder90Day
                oTextBox = txt90DayRemind
                oNextTextBox = txtOEContact

            End If

            oTextBox.Text = strContactName
            oTextBox.TextDataID = intContactID

            oContact.ID = intContactID
            oContact.FullName = strContactName
            'oContact.Order.DefMethod = IIf(optEmail.Checked, ContactMethod.Email, ContactMethod.Fax)
            oContact.OrderContact.PrimaryContact = ContactInfo.Primary ' IIf(optPrimary.Checked, ContactInfo.Primary, ContactInfo.Secondary)

            optEmail.Checked = (oContact.OrderContact.DefMethod = ContactMethod.Email)
            optFax.Checked = (oContact.OrderContact.DefMethod = ContactMethod.Fax)

            oNextTextBox.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvContacts_KeyPress(sender As Object, e As KeyPressEventArgs) Handles dgvContacts.KeyPress

        'Allow user to press key and highlight contact
        Try

            If dgvContacts.RowCount > 0 Then
                SearchLetterTypedInDataGridView(dgvContacts, "FullName", 0, e.KeyChar)
            Else
                MessageBox.Show("There are no listed contacts to search.")
            End If

        Catch ex As Exception
            Dim stackframe As New Diagnostics.StackFrame(1)
            Throw New Exception("An error occurred in routine, '" & stackframe.GetMethod.ReflectedType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "'." & Environment.NewLine & "  Message was: '" & ex.Message & "'")
        End Try

    End Sub


    Public Shared Sub SearchLetterTypedInDataGridView(ByVal dgv As DataGridView, ByVal columnName As String, ByVal columnPosition As Integer, ByVal letterTyped As Char)

        'Try
        '    Dim dt As DataTable = dgv.DataSource
        '    Dim letter As Char = letterTyped
        '    Dim dv As DataView = New DataView(dt)
        '    Dim hasCount As Boolean = False

        '    While (Not hasCount)
        '        dv.Sort = columnName
        '        dv.RowFilter = columnName & " like '" & letter & "%'"
        '        If dv.Count > 0 Then
        '            hasCount = True
        '            Dim x As String = dv(0)(columnPosition).ToString()
        '            Dim bs As New BindingSource
        '            bs.DataSource = dt
        '            dgv.BindingContext(bs).Position = bs.Find(columnName, x)
        '            dgv.CurrentCell = dgv(0, bs.Position)
        '            dgv.CurrentCell.Style.BackColor = Color.Yellow
        '        Else
        '            letter = Chr(Asc(letter) + 1)
        '        End If
        '    End While
        'Catch ex As Exception
        '    Dim stackframe As New Diagnostics.StackFrame(1)
        '    Throw New Exception("An error occurred in search routine, '" & stackframe.GetMethod.ReflectedType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "'." & Environment.NewLine & "  Message was: '" & ex.Message & "'")
        'End Try

    End Sub

End Class
