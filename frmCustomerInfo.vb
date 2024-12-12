Public Class frmCustomerInfo

    Private m_Cus_No As String
    Private m_Customer As cCustomer

    Public Sub New(ByVal pstrCus_No As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        m_Cus_No = pstrCus_No

        Call DataLoad()

    End Sub

    Private Sub DataLoad()

        Try

            m_Customer = New cCustomer(m_Cus_No)

            With m_Customer
                txtCmp_Code.Text = .cmp_code

                If Trim(.cmp_parent.ToString) <> "" Then
                    txtCmp_Parent.Text = GetElementDescription("Cicmpy", "Cmp_Parent", .cmp_parent.Replace("{", "").Replace("}", "")) ' .cmp_parent
                    txtCmp_Parent_Desc.Text = GetElementDescription("Cicmpy", "Cmp_Code", txtCmp_Parent.Text)
                End If

                txtCmp_Name.Text = .cmp_name
                txtCmp_FAdd1.Text = .cmp_fadd1
                txtCmp_FAdd2.Text = .cmp_fadd2
                txtCmp_FAdd3.Text = .cmp_fadd3
                txtCmp_FPC.Text = .cmp_fpc
                txtCmp_FCity.Text = .cmp_fcity
                'txtcmp_fcounty.text = .cmp_fcounty
                txtStateCode.Text = .StateCode
                txtStateCode_Desc.Text = GetElementDescription("AddressStates", "", .cmp_fctry, .StateCode)

                txtCmp_FCtry.Text = .cmp_fctry
                txtCmp_FCtry_Desc.Text = GetElementDescription("Land", "", .cmp_fctry)

                txtCmp_E_Mail.Text = .cmp_e_mail
                txtCmp_Web.Text = .cmp_web
                txtCmp_Fax.Text = .cmp_fax
                txtCmp_Tel.Text = .cmp_tel
                txtCmp_Acc_Man.Text = .cmp_acc_man
                txtCmp_Type.Text = .cmp_type
                txtCmp_Type_Desc.Text = GetElementDescription("cicmpy", "cmp_type", .cmp_type)

                txtCmp_Status.Text = .cmp_status
                txtCmp_Status_Desc.Text = GetElementDescription("cicmpy", "cmp_status", .cmp_status)

                txtDateField1.Text = .datefield1
                txtDateField2.Text = .datefield2
                txtDateField3.Text = .datefield3
                txtDateField4.Text = .datefield4
                txtDateField5.Text = .datefield5
                txtNumberField1.Text = .numberfield1
                txtNumberField2.Text = .numberfield2
                txtNumberField3.Text = .numberfield3
                txtNumberField4.Text = .numberfield4
                txtNumberField5.Text = .numberfield5
                chkYesNoField1.Checked = IIf(.YesNofield1 = 1, True, False)
                chkYesNoField2.Checked = IIf(.YesNofield2 = 1, True, False)
                chkYesNoField3.Checked = IIf(.YesNofield3 = 1, True, False)
                chkYesNoField4.Checked = IIf(.YesNofield4 = 1, True, False)
                chkYesNoField5.Checked = IIf(.YesNofield5 = 1, True, False)
                txtTextField1.Text = .textfield1
                txtTextField2.Text = .textfield2
                txtTextField3.Text = .textfield3
                txtTextField4.Text = .textfield4
                txtTextField5.Text = .textfield5

                txtClassificationID.Text = .ClassificationId
                txtClassificationID_Desc.Text = GetElementDescription("Classifications", "", .ClassificationId)

                txtType_Since.Text = .type_since
                txtStatus_Since.Text = .status_since
                txtDeliveryAddress.Text = .DeliveryAddress
                txtCurrency.Text = .Currency
                txtBankAccountNumber.Text = .BankAccountNumber

                txtPaymentMethod.Text = .PaymentMethod
                txtPaymentMethod_Desc.Text = GetElementDescription("cicmpy", "PaymentMethod", .PaymentMethod)

                txtInvoiceDebtor.Text = .InvoiceDebtor
                txtInvoiceDebtor_Desc.Text = GetElementDescription("Cicmpy", "Cmp_Code", .InvoiceDebtor)

                txtCreditLine.Text = .CreditLine
                txtDiscount.Text = .Discount
                txtDateLastReminder.Text = .DateLastReminder
                chkReminder.Checked = IIf(.Reminder = 1, True, False)
                txtPaymentCondition.Text = .PaymentCondition
                txtEncryptedCreditCard.Text = .CreditCard
                txtExpiryDate.Text = .ExpiryDate
                txtTextField6.Text = .TextField6
                txtTextField7.Text = .TextField7
                txtTextField8.Text = .TextField8
                txtTextField9.Text = .TextField9
                txtTextField10.Text = .TextField10
                txtAccountEmployeeId.Text = .AccountEmployeeId
                chkStatementPrint.Checked = IIf(.StatementPrint = 1, True, False)
                txtStatementNumber.Text = .StatementNumber
                txtStatementDate.Text = .StatementDate
                txtRegion.Text = .region
                txtSalesPersonNumber.Text = .SalesPersonNumber
                txtAccountTypeCode.Text = .AccountTypeCode
                txtAccountRating.Text = .AccountRating
                txtShipVia.Text = .ShipVia
                txtUPSZone.Text = .UPSZone
                chkIsTaxable.Checked = IIf(.IsTaxable = 1, True, False)
                txtTaxCode.Text = .TaxCode
                txtTaxCode2.Text = .TaxCode2
                txtTaxCode3.Text = .TaxCode3
                txtExemptNumber.Text = .ExemptNumber
                chkAllowSubstituteItem.Checked = IIf(.AllowSubstituteItem = 1, True, False)
                chkAllowBackOrders.Checked = IIf(.AllowBackOrders = 1, True, False)
                chkAllowPartialShipment.Checked = IIf(.AllowPartialShipment = 1, True, False)
                txtComment1.Text = .Comment1
                txtComment2.Text = .Comment2
                txtTaxSchedule.Text = .TaxSchedule
                txtCreditCardDescription.Text = .CreditCardDescription
                txtCreditCardHolder.Text = .CreditCardHolder
                txtDefaultInvoiceForm.Text = .DefaultInvoiceForm
                txtLocationShort.Text = .LocationShort
                txtTaxID.Text = .TaxID
                txtEncryptedCreditCard.Text = .EncryptedCreditCard
                txtDeliveryAddress.Text = .DefaultDeliveryAddress

                txtTitle_Cmp_Name.Text = .cmp_code & " - " & .cmp_name

                Call Load_Main_Contact()

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Me.Close()

    End Sub

    Private Sub Load_Addresses()

        Dim strsql As String =
        "SELECT     A.ID, A.Main, T.IsUserDefined, T.TermID, T.Description, " &
        "           C.FullName, A.AddressLine1, A.City, A.TextField1 " &
        "FROM       Addresses A WITH (Nolock) " &
        "INNER JOIN AddressTypes T WITH (Nolock) ON T.ID=A.Type  " &
        "INNER JOIN Cicntp C WITH (Nolock) ON C.cnt_id=A.ContactPerson  " &
        "WHERE  ISNULL(C.active_y,0) = 1 And A.Account='" & m_Customer.cmp_wwn & "' " &
        "ORDER BY   A.Main DESC "

        '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 
        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strsql)

        With dgvAddresses.Columns
            .Add(DGVTextBoxColumn("ID", "ID", 0))
            .Add(DGVCheckBoxColumn("Main", "Main", 50))
            .Add(DGVTextBoxColumn("IsUserDefined", "IsUserDefined", 0))
            .Add(DGVTextBoxColumn("TermID", "TermID", 0))
            .Add(DGVTextBoxColumn("Description", "Description", 100))
            .Add(DGVTextBoxColumn("FullName", "FullName", 160))
            .Add(DGVTextBoxColumn("AddressLine1", "AddressLine1", 160))
            .Add(DGVTextBoxColumn("City", "City", 160))
            .Add(DGVTextBoxColumn("TextField1", "TextField1", 160))
        End With

        dgvAddresses.Columns(Addresses.ID).Visible = False
        dgvAddresses.Columns(Addresses.IsUserDefined).Visible = False
        dgvAddresses.Columns(Addresses.TermID).Visible = False

        dgvAddresses.DataSource = dt
        dgvAddresses.AllowUserToAddRows = False

    End Sub

    Private Sub Load_Main_Contact()

        Dim strsql As String =
        "SELECT     CT.Cnt_ID, CT.Cmp_WWN, ISNULL(CT.FullName, '') AS FullName, " &
        "           ISNULL(CT.TaalCode, '') AS TaalCode, " &
        "           CONVERT(tinyint, CASE WHEN CT.Cnt_Id = C.Cnt_Id THEN 1 ELSE 0 END) AS Main,  " &
        "           ISNULL(CT.Cnt_F_Tel, '') AS Cnt_F_Tel, " &
        "           ISNULL(CT.Cnt_Email, '') AS Cnt_Email, " &
        "           ISNULL(CT.Cnt_Job_Desc, '') AS Cnt_Job_Desc " &
        "FROM       cicntp CT WITH (Nolock) " &
        "INNER JOIN Pred P WITH (Nolock) on P.predcode = CT.predcode  " &
        "INNER JOIN cicmpy C WITH (Nolock) ON CT.cmp_wwn=C.cmp_wwn  " &
        "WHERE ISNULL(CT.active_y,0) = 1 And CT.cmp_wwn = '" & m_Customer.cmp_wwn & "' AND " &
        "           CASE WHEN CT.Cnt_Id = C.Cnt_Id THEN 1 ELSE 0 END = 1 " &
        "ORDER BY   CT.fullname "

        '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strsql)

        If dt.Rows.Count <> 0 Then

            txtCnt_FullName.Text = dt.Rows(0).Item("FullName").ToString
            txtCnt_Job_Description.Text = dt.Rows(0).Item("Cnt_Job_Desc").ToString
            txtCnt_Phone.Text = dt.Rows(0).Item("Cnt_F_Tel").ToString
            txtCnt_Email.Text = dt.Rows(0).Item("Cnt_Email").ToString
            txtCnt_Language.Text = dt.Rows(0).Item("TaalCode").ToString
            txtCnt_Language_Desc.Text = GetElementDescription("Taal", "", txtCnt_Language.Text)

        End If

    End Sub

    Private Sub Load_Contacts()

        Dim strsql As String =
        "SELECT     CT.Cnt_ID, CT.Cmp_WWN, CT.FullName, " &
        "           CONVERT(tinyint, CASE WHEN CT.Cnt_Id = C.Cnt_Id THEN 1 ELSE 0 END) AS Main,  " &
        "           CT.Active_Y, P.Abbreviation, CT.Cnt_F_Tel, CT.Cnt_Job_Desc  " &
        "FROM       cicntp CT WITH (Nolock) " &
        "INNER JOIN Pred P WITH (Nolock) on P.predcode = CT.predcode  " &
        "INNER JOIN cicmpy C WITH (Nolock) ON CT.cmp_wwn=C.cmp_wwn  " &
        "WHERE  ISNULL(CT.active_y,0) = 1 And CT.cmp_wwn = '" & m_Customer.cmp_wwn & "'  " &
        "ORDER BY   CT.fullname "

        '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strsql)

        With dgvContacts.Columns
            .Add(DGVTextBoxColumn("Cnt_ID", "Cnt_ID", 0))
            .Add(DGVTextBoxColumn("Cmp_WWN", "Cmp_WWN", 0))
            .Add(DGVTextBoxColumn("FullName", "Name", 200))
            .Add(DGVCheckBoxColumn("Main", "Main", 40))
            .Add(DGVCheckBoxColumn("Active_Y", "Active", 40))
            .Add(DGVTextBoxColumn("Abbreviation", "Title", 100))
            .Add(DGVTextBoxColumn("Cnt_F_Tel", "Phone", 160))
            .Add(DGVTextBoxColumn("Cnt_Job_Desc", "Job Description", 250))

        End With

        dgvContacts.Columns(Contacts.Cnt_ID).Visible = False
        dgvContacts.Columns(Contacts.Cmp_WWN).Visible = False

        dgvContacts.DataSource = dt
        dgvContacts.AllowUserToAddRows = False

    End Sub

    Private Sub TabPage2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage2.Enter

        If dgvAddresses.Rows.Count = 0 Then Call Load_Addresses()

    End Sub

    Private Sub TabPage3_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage3.Enter

        If dgvContacts.Rows.Count = 0 Then Call Load_Contacts()

    End Sub

    Private Enum Addresses

        ID
        Main
        IsUserDefined
        TermID
        Description
        FullName
        AddressLine1
        City
        TextField1

    End Enum

    Private Enum Contacts

        Cnt_ID
        Cmp_WWN
        FullName
        Main
        Active_Y
        Abbreviation
        Cnt_F_Tel
        Cnt_Job_Desc

    End Enum

    Private Sub TabControl1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TabControl1.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose.PerformClick()
        End Select

    End Sub

End Class