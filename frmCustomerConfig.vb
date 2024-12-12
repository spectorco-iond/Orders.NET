Public Class frmCustomerConfig

    Private m_LastTab As String

    Private m_strCus_No As String = String.Empty
    Private db As New cDBA

    Private m_oGeneral As COEFlash
    Private m_dtGeneral As DataTable = New DataTable

    Private m_oOEFlash As COEFlash
    Private m_dtOEFlash As DataTable = New DataTable

    'Private m_oCharges As cOEFlashCharges
    'Private m_dtCharges As DataTable = New DataTable

    Private m_oCorrespondance As cOEFlashCorrespondance
    Private m_dtCorrespondance As DataTable = New DataTable

    Private m_oCustomer As cOEFlashCustomer
    Private m_dtCustomer As DataTable = New DataTable

    'Private m_oPricing As cOEFlashPricing
    'Private m_dtPricing As DataTable = New DataTable

    Private m_oShipping As cOEFlashShipping
    Private m_dtShipping As DataTable = New DataTable

    Private blnLoading As Boolean = False
    Private blnDeleting As Boolean = False
    Private blnInserting As Boolean = False

    Delegate Sub SetColumnIndex(ByRef dgv As DataGridView, ByVal i As Integer)

#Region "Public constructors ##############################################"

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'Call InitGrids()

    End Sub

    Public Sub New(ByVal pstrCus_No As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'Call InitGrids()

        m_strCus_No = pstrCus_No
        txtCusNo.Text = m_strCus_No

        Call LoadCustomerConfig()
        'Call LoadGrids()

    End Sub

#End Region

#Region "Private grid procedures ##########################################"

    'Private Sub InitGrids()

    '    'InitOEFlashGrid()
    '    'InitGeneralGrid()
    '    ''InitChargesGrid()
    '    ''InitPricingGrid()
    '    'InitShippingGrid()
    '    'InitCorrespondanceGrid()

    'End Sub

    'Private Sub InitOEFlashGrid()

    '    Try

    '        With dgvOEFlash.Columns

    '            .Add(DGVTextBoxColumn("ID", "ID", 0))
    '            .Add(DGVTextBoxColumn("Comments", "Comments", 600))
    '            .Add(DGVTextBoxColumn("Cus_Sort", "Cus_Sort", 0))

    '        End With

    '        dgvOEFlash.Columns(OEFlashColumns.ID).Visible = False
    '        dgvOEFlash.Columns(OEFlashColumns.Cus_Sort).Visible = False

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try


    'End Sub

    'Private Sub InitGeneralGrid()

    '    Try

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub InitChargesGrid()

    '    Try

    '        With dgvCharges.Columns

    '            .Add(DGVTextBoxColumn("Charge_Usage_ID", "Charge_Usage_ID", 0))
    '            .Add(DGVTextBoxColumn("Charge_CD", "Charge", 115))
    '            .Add(DGVTextBoxColumn("Cus_No", "Cus_No", 0))
    '            .Add(DGVCheckBoxColumn("Always_Use", "Always", 42))
    '            .Add(DGVCheckBoxColumn("Never_Use", "Never", 35))
    '            .Add(DGVCheckBoxColumn("No_Charge", "N/C", 25))
    '            .Add(DGVCalendarColumn("Charge_From", "From", 70))
    '            .Add(DGVCalendarColumn("Charge_To", "To", 70))
    '            .Add(DGVCheckBoxColumn("Blind", "Blind", 27))
    '            .Add(DGVTextBoxColumn("Apply_To_Ship_To", "On ShipTo", 50))
    '            .Add(DGVTextBoxColumn("Apply_To_Program", "On Logo", 50))
    '            .Add(DGVTextBoxColumn("When_Qty_LT", "Qty GT", 40))
    '            .Add(DGVTextBoxColumn("When_Qty_GT", "Qty LT", 40))
    '            .Add(DGVTextBoxColumn("When_Amt_LT", "Amt GT", 40))
    '            .Add(DGVTextBoxColumn("When_Amt_GT", "Amt LT", 40))
    '            .Add(DGVCheckBoxColumn("When_Req", "When Req", 32))
    '            .Add(DGVTextBoxColumn("Autorized_By", "Autorized By", 75))
    '            .Add(DGVTextBoxColumn("Comments", "Comments", 200))

    '        End With

    '        dgvCharges.Columns(ChargesColumns.Charge_Usage_ID).Visible = False
    '        dgvCharges.Columns(ChargesColumns.Cus_No).Visible = False
    '        dgvCharges.Columns(ChargesColumns.Apply_To_Ship_To).Visible = False
    '        dgvCharges.Columns(ChargesColumns.Apply_To_Program).Visible = False

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub InitPricingGrid()

    '    Try

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub InitShippingGrid()

    '    Try

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub InitCorrespondanceGrid()

    '    Try

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub LoadGrids()

    '    LoadOEFlashGrid()
    '    LoadGeneralGrid()
    '    'LoadChargesGrid()
    '    LoadPricingGrid()
    '    LoadShippingGrid()
    '    LoadCorrespondanceGrid()

    'End Sub

    'Private Sub LoadOEFlashGrid()

    '    Dim strSql As String 

    '    Try
    '        strSql = _
    '        "SELECT ID, Cmt AS Comments, Cus_Sort " & _
    '        "FROM   Exact_Custom_OeFlash_CustomerComments WITH (Nolock) " & _
    '        "WHERE  Cus_no = '" & m_strCus_No & "' ORDER BY Cus_Sort"

    '        m_dtOEFlash = db.DataTable(strSql)
    '        dgvOEFlash.DataSource = m_dtOEFlash

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub LoadGeneralGrid()

    '    'Dim strSql As String

    '    'Try
    '    '    strSql = _
    '    '    "SELECT ID, Cmt AS Comments, Cus_Sort " & _
    '    '    "FROM   Exact_Custom_OeFlash_CustomerComments WITH (Nolock) " & _
    '    '    "WHERE  Cus_no = '" & m_strCus_No & "' ORDER BY Cus_Sort"

    '    '    Dim m_dtgeneral As DataTable = db.DataTable(strSql)
    '    '    dgvGeneral.DataSource = m_dtgeneral

    '    'Catch er As Exception
    '    '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    'End Try

    'End Sub

    'Private Sub LoadChargesGrid()

    '    Dim strSql As String

    '    Try
    '        strSql = _
    '        "SELECT		CU.CHARGE_USAGE_ID, CU.Charge_CD, CU.Cus_No, CU.Always_Use, " & _
    '        "           CU.Never_Use, CU.No_Charge, ISNULL(CU.Charge_From, GETDATE()) AS Charge_From, " & _
    '        "           ISNULL(CU.Charge_To, GETDATE()) AS Charge_To, CU.Blind, CU.Apply_To_Ship_To, CU.Apply_To_Program, " & _
    '        "           CU.When_Qty_GT, CU.When_Qty_LT, CU.When_Amt_GT, CU.When_Amt_LT, " & _
    '        "			CU.When_Req, CU.Autorized_By, CU.Comments " & _
    '        "FROM		MDB_CFG_CHARGE_USAGE CU WITH (Nolock) " & _
    '        "INNER JOIN	MDB_CFG_CHARGE C WITH (Nolock) ON CU.Charge_CD = C.Charge_CD " & _
    '        "WHERE		CU.Cus_No = '" & m_strCus_No & "' "
    '        If Not chkChargesAllDates.Checked Then strSql &= " AND " & _
    '        "           ( " & _
    '        "				(CU.Charge_From IS NULL AND CU.Charge_To IS NULL) OR " & _
    '        "				(CU.Charge_From <= GETDATE()) AND CU.Charge_To >= DATEDIFF(d, 0, GETDATE()) OR " & _
    '        "				(CU.Charge_From <= GETDATE()) AND (CU.Charge_To IS NULL) OR " & _
    '        "				(CU.Charge_From IS NULL AND CU.Charge_To >= DATEDIFF(d, 0, GETDATE()))" & _
    '        "			) "
    '        strSql &= " ORDER BY CU.Charge_CD, CU.Charge_To DESC, CU.Charge_From "

    '        m_dtCharges = db.DataTable(strSql)
    '        dgvCharges.DataSource = m_dtCharges

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub LoadPricingGrid()

    '    'Dim strSql As String

    '    'Try
    '    '    strSql = _
    '    '    "SELECT ID, Cmt AS Comments, Cus_Sort " & _
    '    '    "FROM   Exact_Custom_OeFlash_CustomerComments WITH (Nolock) " & _
    '    '    "WHERE  Cus_no = '" & m_strCus_No & "' ORDER BY Cus_Sort"

    '    '    Dim m_dtpricing As DataTable = db.DataTable(strSql)
    '    '    dgvPricing.DataSource = m_dtpricing

    '    'Catch er As Exception
    '    '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    'End Try

    'End Sub

    'Private Sub LoadShippingGrid()

    '    'Dim strSql As String

    '    'Try
    '    '    strSql = _
    '    '    "SELECT ID, Cmt AS Comments, Cus_Sort " & _
    '    '    "FROM   Exact_Custom_OeFlash_CustomerComments WITH (Nolock) " & _
    '    '    "WHERE  Cus_no = '" & m_strCus_No & "' ORDER BY Cus_Sort"

    '    '    Dim m_dtshipping As DataTable = db.DataTable(strSql)
    '    '    dgvShipping.DataSource = m_dtshipping

    '    'Catch er As Exception
    '    '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    'End Try

    'End Sub

    'Private Sub LoadCorrespondanceGrid()

    '    'Dim strSql As String

    '    'Try
    '    '    strSql = _
    '    '    "SELECT ID, Cmt AS Comments, Cus_Sort " & _
    '    '    "FROM   Exact_Custom_OeFlash_CustomerComments WITH (Nolock) " & _
    '    '    "WHERE  Cus_no = '" & m_strCus_No & "' ORDER BY Cus_Sort"

    '    '    Dim m_dtcorrespondance As DataTable = db.DataTable(strSql)
    '    '    dgvCorrespondance.DataSource = m_dtcorrespondance

    '    'Catch er As Exception
    '    '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    'End Try

    'End Sub

    Private Sub LoadCustomerConfig()

        Try

            Dim strSql As String = _
            "SELECT Cmp_Name, ISNULL(ClassificationID, '') AS ClassificationID, ISNULL(TextField1, '') AS EQP  " & _
            "FROM   cicmpy WITH (Nolock) " & _
            "WHERE  cmp_code = '" & txtCusNo.Text & "'"

            Dim dt As New DataTable
            dt = db.DataTable(strSql)
            If dt.Rows.Count > 0 Then
                Dim dtRow As DataRow = dt.Rows(0)
                txtCusName.Text = dtRow.Item("Cmp_Name").ToString
                txtClassID.Text = "Star Lvl: " & dtRow.Item("ClassificationID").ToString
                If Trim(dtRow.Item("EQP")) = "" Then
                    txtEQP.Text = ""
                Else
                    txtEQP.Text = "Terms: " & dtRow.Item("EQP")
                End If

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "PUBLIC PREPORTIES"

    Public Property Cus_No() As String
        Get
            Cus_No = m_strCus_No
        End Get
        Set(ByVal value As String)
            m_strCus_No = value

            'Call InitGrids()
            Call LoadCustomerConfig()
            'Call LoadGrids()

        End Set
    End Property

#End Region

    Private Sub chkOEFlash_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOEFlash.CheckedChanged

        dgvOEFlash.Visible = chkOEFlash.Checked
        dgvGeneral.Visible = Not (chkOEFlash.Checked)

    End Sub

    'Private Sub dgvCharges_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    '    Try

    '        If blnDeleting Then Exit Sub

    '        If dgvCharges.Rows.Count = 0 Then Exit Sub
    '        'If dgvItems.CurrentRow.Index = 0 Then Exit Sub

    '        If e.ColumnIndex < ChargesColumns.Charge_CD Then

    '            Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
    '            dgvCharges.BeginInvoke(oInvoke, dgvCharges, ChargesColumns.Charge_CD)

    '        ElseIf e.ColumnIndex = ChargesColumns.Comments Then ' LastCol

    '            Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
    '            dgvCharges.BeginInvoke(oInvoke, dgvCharges, ChargesColumns.Charge_CD)

    '            'ElseIf e.ColumnIndex > Columns.Item_No And IIf(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Equals(DBNull.Value), "", dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value) = "" Then
    '        ElseIf e.ColumnIndex > ChargesColumns.Charge_CD And dgvCharges.Rows(e.RowIndex).Cells(ChargesColumns.Charge_CD).Value Is Nothing Then
    '            'ElseIf e.ColumnIndex > Columns.Item_No And IIf(dgvItems.CurrentRow.Cells(Columns.Item_No).Equals(DBNull.Value), "", dgvItems.CurrentRow.Cells(Columns.Item_No).Value) = "" Then

    '            Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
    '            dgvCharges.BeginInvoke(oInvoke, dgvCharges, ChargesColumns.Charge_CD)

    '        ElseIf e.ColumnIndex > ChargesColumns.Charge_CD And Trim(dgvCharges.Rows(e.RowIndex).Cells(ChargesColumns.Charge_CD).Value.ToString) = "" Then

    '            Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
    '            dgvCharges.BeginInvoke(oInvoke, dgvCharges, ChargesColumns.Charge_CD)

    '        End If

    '        'If dgvItems.CurrentCell.ColumnIndex > Columns.Item_No And dgvItems.CurrentCell.ColumnIndex <= Columns.Item_Desc_2 And Not (dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value Is Nothing) Then
    '        '    dgvItems.BeginEdit(True)
    '        'End If

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub dgvCharges_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs)

    '    If blnLoading Then Exit Sub

    '    If m_oCharges Is Nothing Then
    '        m_oCharges = New cOEFlashCharges(m_oOrder.Ordhead.Cus_No) ' THAT CODE SHOULD NEVER HAPPEN
    '    End If

    '    'If e.FormattedValue.ToString = "" Then Exit Sub

    '    Try
    '        Select Case e.ColumnIndex

    '            Case ChargesColumns.Charge_CD
    '                m_oCharges.Charge_CD = Trim(e.FormattedValue).ToString

    '                'Case ChargesColumns.Cus_No
    '                '    m_oCharges.Cus_No = Trim(e.FormattedValue).ToString

    '            Case ChargesColumns.Always_Use
    '                m_oCharges.Always_Use = (e.FormattedValue)

    '            Case ChargesColumns.Never_Use
    '                m_oCharges.Never_Use = (e.FormattedValue)

    '            Case ChargesColumns.No_Charge
    '                m_oCharges.No_Charge = (e.FormattedValue)

    '            Case ChargesColumns.Charge_From
    '                If Trim(e.FormattedValue) <> "" Then
    '                    If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format, True)
    '                    m_oCharges.Charge_From = CDate(e.FormattedValue)
    '                End If

    '            Case ChargesColumns.Charge_To
    '                If Trim(e.FormattedValue) <> "" Then
    '                    If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format, True)
    '                    m_oCharges.Charge_To = CDate(e.FormattedValue)
    '                End If

    '            Case ChargesColumns.Blind
    '                m_oCharges.Blind = (e.FormattedValue)

    '            Case ChargesColumns.Apply_To_Ship_To
    '                m_oCharges.Apply_To_Ship_To = Trim(e.FormattedValue).ToString

    '            Case ChargesColumns.Apply_To_Program
    '                m_oCharges.Apply_To_Program = Trim(e.FormattedValue).ToString

    '            Case ChargesColumns.When_Qty_GT
    '                If Trim(e.FormattedValue) <> "" Then
    '                    If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
    '                    m_oCharges.When_Qty_GT = (e.FormattedValue)
    '                End If

    '            Case ChargesColumns.When_Qty_LT
    '                If Trim(e.FormattedValue) <> "" Then
    '                    If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
    '                    m_oCharges.When_Qty_LT = (e.FormattedValue)
    '                End If

    '            Case ChargesColumns.When_Amt_GT
    '                If Trim(e.FormattedValue) <> "" Then
    '                    If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
    '                    m_oCharges.When_Amt_GT = (e.FormattedValue)
    '                End If

    '            Case ChargesColumns.When_Amt_LT
    '                If Trim(e.FormattedValue) <> "" Then
    '                    If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
    '                    m_oCharges.When_Amt_LT = (e.FormattedValue)
    '                End If

    '            Case ChargesColumns.When_Req
    '                m_oCharges.When_Req = (e.FormattedValue)

    '            Case ChargesColumns.Autorized_By
    '                m_oCharges.Autorized_By = Trim(e.FormattedValue).ToString

    '            Case ChargesColumns.Comments
    '                m_oCharges.Comments = Trim(e.FormattedValue).ToString

    '        End Select

    '    Catch oe_er As OEException
    '        e.Cancel = oe_er.Cancel
    '        If oe_er.ShowMessage Then MsgBox(oe_er.Message)
    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try


    'End Sub

    'Private Sub dgvCharges_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs)
    '    Debug.Print(e.Column.Width)
    'End Sub

    'Private Sub dgvCharges_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

    '    Try
    '        Select Case dgvCharges.CurrentCell.ColumnIndex

    '            Case ChargesColumns.Charge_CD

    '                If e.KeyCode = Keys.F7 Then

    '                    Dim oSearch As New cSearch("MDB-CFG_CHARGE")

    '                    If oSearch.Form.FoundRow <> -1 Then
    '                        dgvCharges.CurrentCell.Value = oSearch.Form.FoundElementValue

    '                    End If
    '                    oSearch = Nothing

    '                End If

    '        End Select

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub dgvCharges_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)

    '    Try
    '        If m_oCharges Is Nothing Then Exit Sub
    '        If m_oCharges.Dirty Then m_oCharges.Save()

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub dgvCharges_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    '    Try
    '        With dgvCharges.Rows(e.RowIndex)
    '            cmdChargeDelete.Enabled = Not (.Cells(ChargesColumns.Charge_Usage_ID).Value Is Nothing)

    '            If .Cells(ChargesColumns.Charge_Usage_ID).Value.Equals(DBNull.Value) Then
    '                m_oCharges = New cOEFlashCharges(m_oOrder.Ordhead.Cus_No) ' m_oOrder)
    '            Else
    '                m_oCharges = New cOEFlashCharges(CInt(.Cells(ChargesColumns.Charge_Usage_ID).Value))
    '            End If

    '        End With

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub dgvCharges_RowLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    '    Try

    '        If Not m_oCharges Is Nothing Then

    '            m_oCharges.Save()

    '            With dgvCharges.CurrentRow.Cells(ChargesColumns.Charge_Usage_ID)
    '                If .Value.Equals(DBNull.Value) Then
    '                    .Value = m_oCharges.Charge_Usage_ID
    '                End If
    '            End With

    '        End If

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub SetColumnIndexSub(ByRef pdgv As DataGridView, ByVal columnIndex As Integer)

    '    pdgv.CurrentCell = pdgv.CurrentRow.Cells(columnIndex)

    'End Sub

    'Private Sub LineInsert(ByRef pdgv As DataGridView, ByRef pdt As DataTable, ByVal pintCell As Integer)

    '    Dim drNewRow As DataRow
    '    pdgv.AllowUserToAddRows = True

    '    drNewRow = pdt.NewRow

    '    pdt.Rows.Add(drNewRow)

    '    pdgv.AllowUserToAddRows = False
    '    pdgv.CurrentCell = pdgv.Rows(pdgv.Rows.Count - 1).Cells(pintCell) ' Columns.Item_No

    '    pdgv.Focus()

    'End Sub

    'Private Sub cmdChargeAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Try

    '        blnInserting = True
    '        'Call ProductLineInsert()

    '        If dgvCharges.Rows.Count = 0 Then
    '            Call LineInsert(dgvCharges, m_dtCharges, ChargesColumns.Charge_CD)
    '        ElseIf Not (dgvCharges.Rows(dgvCharges.Rows.Count - 1).Cells(ChargesColumns.Charge_CD).Value.Equals(DBNull.Value)) Then
    '            If Not Trim(dgvCharges.Rows(dgvCharges.Rows.Count - 1).Cells(ChargesColumns.Charge_CD).Value).Equals("") Then
    '                Call LineInsert(dgvCharges, m_dtCharges, ChargesColumns.Charge_CD)
    '            Else
    '                dgvCharges.CurrentCell = dgvCharges.Rows(dgvCharges.Rows.Count - 1).Cells(ChargesColumns.Charge_CD)
    '            End If
    '        Else
    '            dgvCharges.CurrentCell = dgvCharges.Rows(dgvCharges.Rows.Count - 1).Cells(ChargesColumns.Charge_CD)
    '        End If

    '        blnInserting = False

    '        'g_oOrdline = New cOrdLine(m_oOrder)

    '    Catch oe_er As OEException
    '        MsgBox(oe_er.Message)
    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub cmdChargeDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Try
    '        dgvCharges.CurrentCell = dgvCharges.CurrentRow.Cells(ChargesColumns.Charge_CD)

    '        Call ChargeLineDelete()

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    'Private Sub ChargeLineDelete(Optional ByVal pblnAsk As Boolean = True)

    '    Try
    '        If dgvCharges.Rows.Count = 0 Then Exit Sub
    '        If dgvCharges.Rows(0).Cells(ChargesColumns.Charge_CD).Value.ToString = "" Then Exit Sub

    '        ' ask if it is ok to delete. If not OK, then exit sub right away
    '        If pblnAsk Then
    '            If Trim(dgvCharges.CurrentRow.Cells(ChargesColumns.Charge_CD).Value.ToString) <> "" Then
    '                Dim mbrOkToDelete As MsgBoxResult
    '                mbrOkToDelete = MsgBox("OK to delete ?", MsgBoxStyle.YesNo, "Order entry interface")

    '                If mbrOkToDelete = MsgBoxResult.No Then Exit Sub
    '            End If
    '        End If

    '        Dim iPos As Integer = dgvCharges.CurrentRow.Index

    '        If dgvCharges.CurrentRow.Index >= 0 Then
    '            m_oCharges.Delete()
    '            dgvCharges.Rows.RemoveAt(iPos)
    '        End If

    '        If dgvCharges.Rows.Count > 0 Then
    '            If iPos >= dgvCharges.Rows.Count Then
    '                dgvCharges.CurrentCell = dgvCharges.Rows(0).Cells(ChargesColumns.Charge_CD)
    '                dgvCharges.CurrentCell = dgvCharges.Rows(iPos - 1).Cells(ChargesColumns.Charge_CD)
    '            Else
    '                If iPos < 0 Then iPos = 0
    '                dgvCharges.CurrentCell = dgvCharges.Rows(iPos).Cells(ChargesColumns.Charge_CD)
    '            End If
    '        End If

    '        dgvCharges.Focus()

    '        blnDeleting = False

    '    Catch oe_er As OEException
    '        MsgBox(oe_er.Message)
    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '        blnDeleting = False
    '    End Try

    'End Sub

    'Private Sub chkChargesAllDates_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Call LoadChargesGrid()

    'End Sub

    Private Sub OETab_Charges_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles OETab_Charges.Enter

        uc_OEFlashCharges.Fill(m_oOrder.Ordhead.Cus_No)

    End Sub

    Private Sub OETab_Pricing_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles OETab_Pricing.Enter

        'Uc_OEFlashPricing.Fill(m_oOrder.Ordhead.Cus_No)

    End Sub

    Private Sub tcOEFlash_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tcOEFlash.Click

        Try

            ' First, we save the last tab
            Select Case m_LastTab

                Case OETab_General.Name
                    If tcOEFlash.SelectedIndex <> Tabs.General Then
                        Uc_OEFlashGeneral.Save()
                    End If

                Case OETab_Charges.Name
                    If tcOEFlash.SelectedIndex <> Tabs.Charges Then
                        uc_OEFlashCharges.Save()
                    End If

                Case OETab_Pricing.Name
                    If tcOEFlash.SelectedIndex <> Tabs.Pricing Then
                        'Uc_OEFlashPricing.Save()
                    End If

                Case OETab_Shipping.Name
                    If tcOEFlash.SelectedIndex <> Tabs.Shipping Then
                        Uc_OEFlashShipping.Save()
                    End If

                Case OETab_Correspondance.Name
                    If tcOEFlash.SelectedIndex <> Tabs.Correspondance Then
                        Uc_OEFlashCorrespondance.Save()
                    End If

                Case Else

            End Select

            ' Then we get the new one, for the next time.
            Select Case tcOEFlash.SelectedIndex

                Case Tabs.General
                    m_LastTab = OETab_General.Name
                Case Tabs.Charges
                    m_LastTab = OETab_Charges.Name
                Case Tabs.Pricing
                    m_LastTab = OETab_Pricing.Name
                Case Tabs.Shipping
                    m_LastTab = OETab_Shipping.Name
                Case Tabs.Correspondance
                    m_LastTab = OETab_Correspondance.Name
                Case Else
                    m_LastTab = ""

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Enum Tabs

        General
        Charges
        Pricing
        Shipping
        Correspondance

    End Enum

End Class