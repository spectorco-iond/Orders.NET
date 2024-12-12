Public Class ucOEFlashPricing

    Private m_strCus_No As String = String.Empty
    Private db As New cDBA

    Private m_oPricing As cOEFlashPricing
    Private m_dt As DataTable = New DataTable

    Private blnLoading As Boolean = False
    Private blnDeleting As Boolean = False
    'Private blnInserting As Boolean = False

    Delegate Sub SetColumnIndex(ByRef dgv As DataGridView, ByVal i As Integer)

#Region "Public constructors ##############################################"

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Call InitGrid()

    End Sub

#End Region

#Region "Public procedures ################################################"

    Public Sub Fill(ByVal pstrCus_No As String)

        Cus_No = Trim(pstrCus_No)

    End Sub

    Public Sub Save()

        Try
            If Not blnLoading Then
                If Not m_oPricing Is Nothing Then

                    If m_oPricing.Dirty Then
                        m_oPricing.Save()
                    End If

                    With dgvPricing.CurrentRow.Cells(Columns.OE_Price_ID)
                        If .Value.Equals(DBNull.Value) Then
                            .Value = m_oPricing.OE_Price_ID
                        ElseIf .Value <> m_oPricing.OE_Price_ID Then
                            .Value = m_oPricing.OE_Price_ID
                        End If
                    End With

                End If
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private grid procedures ##########################################"

    Private Sub InitGrid()

        Try
            If dgvPricing.Columns.Count <> 0 Then Exit Sub

            With dgvPricing.Columns

                .Add(DGVTextBoxColumn("OE_Price_ID", "OE_Price_ID", 0))
                .Add(DGVTextBoxColumn("Cus_No", "Cus_No", 0))
                .Add(DGVTextBoxColumn("Program_Name", "Program", 120))
                .Add(DGVTextBoxColumn("Item_CD", "Item CD", 50))
                .Add(DGVTextBoxColumn("Prc_Or_Disc_1", "Price", 35))
                Dim cbComboColumnDef As DataGridViewComboBoxColumn
                'dgvItems.Columns.Add(DGVTextBoxColumn("Industry", "Industry", 100))
                cbComboColumnDef = DGVComboBoxColumn("EQP_Type", "EQP Type", 95)
                With cbComboColumnDef
                    .DataSource = m_oOrder.GetEQPTypes()
                    .ValueMember = "EQP_Type"
                    .DisplayMember = "EQP_Type"
                End With
                .Add(cbComboColumnDef)
                .Add(DGVTextBoxColumn("EQP_Plus_Pct", "EQP Pct", 30))
                .Add(DGVTextBoxColumn("Minimum_Qty_1", "Min Qty", 35))
                .Add(DGVCalendarColumn("Start_Date", "From", 70))
                .Add(DGVCalendarColumn("End_Date", "To", 70))
                .Add(DGVCheckBoxColumn("OEI_Warning", "OEI Warn", 32))
                .Add(DGVCheckBoxColumn("Macola_Import", "Macola Import", 40))
                .Add(DGVCheckBoxColumn("When_Req", "When Req", 32))
                .Add(DGVTextBoxColumn("Autorized_By", "Autorized By", 75))
                .Add(DGVTextBoxColumn("Comments", "Comments", 200))

            End With

            dgvPricing.Columns(Columns.OE_Price_ID).Visible = False
            dgvPricing.Columns(Columns.Cus_No).Visible = False

            Dim oCellStyle As New DataGridViewCellStyle()

            oCellStyle.Format = "##,##0.00"
            oCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvPricing.Columns(Columns.EQP_Plus_Pct).DefaultCellStyle = oCellStyle
            dgvPricing.Columns(Columns.Minimum_Qty_1).DefaultCellStyle = oCellStyle
            dgvPricing.Columns(Columns.Prc_Or_Disc_1).DefaultCellStyle = oCellStyle

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadGrid()

        Dim strSql As String

        Try
            strSql = _
            "SELECT		P.OE_Price_ID, P.CD_TP_1_Cust_No AS Cus_No, P.Program_Name, ISNULL(P.Item_CD, '') AS Item_CD, " & _
            "           P.Prc_Or_Disc_1, P.EQP_Type, P.EQP_Plus_Pct, P.Minimum_Qty_1, " & _
            "           ISNULL(P.Start_Date, GETDATE()) AS Start_Date, ISNULL(P.End_Date, GETDATE()) AS End_Date, " & _
            "           P.OEI_Warning, P.Macola_Import, P.When_Req, P.Autorized_By, P.Comments " & _
            "FROM		MDB_OE_PRICE P WITH (Nolock) " & _
            "WHERE		P.CD_TP_1_Cust_No = '" & m_strCus_No & "' "
            If Not chkAllDates.Checked Then strSql &= " AND " & _
            "           ( " & _
            "				(P.Start_Date IS NULL AND P.End_Date IS NULL) OR " & _
            "				(P.Start_Date <= GETDATE()) AND P.End_Date >= DATEDIFF(d, 0, GETDATE()) OR " & _
            "				(P.Start_Date <= GETDATE()) AND (P.End_Date IS NULL) OR " & _
            "				(P.Start_Date IS NULL AND P.End_Date >= DATEDIFF(d, 0, GETDATE()))" & _
            "			) "
            strSql &= " ORDER BY P.Program_Name, ISNULL(P.Item_CD, ''), P.End_Date DESC, P.Start_Date "

            blnLoading = True

            m_dt = db.DataTable(strSql)
            dgvPricing.DataSource = m_dt

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        blnLoading = False

    End Sub

#End Region

#Region "PUBLIC PREPORTIES"

    Public Property Cus_No() As String
        Get
            Cus_No = m_strCus_No
        End Get
        Set(ByVal value As String)
            m_strCus_No = value

            Call LoadGrid()

        End Set
    End Property

#End Region

    Private Sub dgvPricing_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPricing.CellEnter

        Try

            If blnDeleting Then Exit Sub

            If dgvPricing.Rows.Count = 0 Then Exit Sub
            'If dgvItems.CurrentRow.Index = 0 Then Exit Sub

            If e.ColumnIndex < Columns.Program_Name Then

                Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
                dgvPricing.BeginInvoke(oInvoke, dgvPricing, Columns.Program_Name)

            ElseIf e.ColumnIndex = Columns.Comments Then ' LastCol

                Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
                dgvPricing.BeginInvoke(oInvoke, dgvPricing, Columns.Program_Name)

                'ElseIf e.ColumnIndex > Columns.Item_No And IIf(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Equals(DBNull.Value), "", dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value) = "" Then
            ElseIf e.ColumnIndex > Columns.Program_Name And dgvPricing.Rows(e.RowIndex).Cells(Columns.Program_Name).Value Is Nothing Then
                'ElseIf e.ColumnIndex > Columns.Item_No And IIf(dgvItems.CurrentRow.Cells(Columns.Item_No).Equals(DBNull.Value), "", dgvItems.CurrentRow.Cells(Columns.Item_No).Value) = "" Then

                Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
                dgvPricing.BeginInvoke(oInvoke, dgvPricing, Columns.Program_Name)

            ElseIf e.ColumnIndex > Columns.Program_Name And Trim(dgvPricing.Rows(e.RowIndex).Cells(Columns.Program_Name).Value.ToString) = "" Then

                Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
                dgvPricing.BeginInvoke(oInvoke, dgvPricing, Columns.Program_Name)

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvPricing_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvPricing.CellValidating

        If blnLoading Then Exit Sub

        If m_oPricing Is Nothing Then
            m_oPricing = New cOEFlashPricing(m_oOrder.Ordhead.Cus_No) ' THAT CODE SHOULD NEVER HAPPEN
        End If

        Try
            Select Case e.ColumnIndex

                Case Columns.Program_Name
                    If Trim(e.FormattedValue).ToString <> "" Then
                        'If m_oPricing.Charge_Exist(Trim(e.FormattedValue).ToString) Then
                        m_oPricing.Program_Name = Trim(e.FormattedValue).ToString
                        'Else
                        '    Throw New OEException(OEError.Charge_Not_Found, True)
                        'End If
                    End If

                Case Columns.Item_CD
                    m_oPricing.Item_CD = (e.FormattedValue)

                Case Columns.Prc_Or_Disc_1
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                        m_oPricing.Prc_Or_Disc_1 = (e.FormattedValue)
                    End If

                Case Columns.EQP_Type
                    m_oPricing.EQP_Type = (e.FormattedValue)

                Case Columns.EQP_Plus_Pct
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                        m_oPricing.EQP_Plus_Pct = (e.FormattedValue)
                    End If

                Case Columns.Minimum_Qty_1
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                        m_oPricing.Minimum_Qty_1 = (e.FormattedValue)
                    End If

                Case Columns.Start_Date
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format, True)
                        m_oPricing.Start_Date = CDate(e.FormattedValue)
                    End If

                Case Columns.End_Date
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format, True)
                        m_oPricing.End_Date = CDate(e.FormattedValue)
                    End If

                Case Columns.OEI_Warning
                    m_oPricing.OEI_Warning = (e.FormattedValue)

                Case Columns.Macola_Import
                    m_oPricing.Macola_Import = (e.FormattedValue)

                Case Columns.When_Req
                    m_oPricing.When_Req = (e.FormattedValue)

                Case Columns.Autorized_By
                    m_oPricing.Autorized_By = Trim(e.FormattedValue).ToString

                Case Columns.Comments
                    m_oPricing.Comments = Trim(e.FormattedValue).ToString

            End Select

        Catch oe_er As OEException
            e.Cancel = oe_er.Cancel
            If oe_er.ShowMessage Then MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try


    End Sub


    Private Sub dgvPricing_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvPricing.KeyDown

        Try
            'Select Case dgvPricing.CurrentCell.ColumnIndex

            '    Case Columns.Program_Name

            '        If e.KeyCode = Keys.F7 Then

            '            Dim oSearch As New cSearch("MDB-CFG-CHARGE")

            '            If oSearch.Form.FoundRow <> -1 Then
            '                dgvPricing.CurrentCell.Value = oSearch.Form.FoundElementValue

            '            End If
            '            oSearch = Nothing

            '        End If

            'End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvPricing_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvPricing.LostFocus

        Try
            If m_oPricing Is Nothing Then Exit Sub

            If m_oPricing.Dirty Then
                m_oPricing.Save()
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvPricing_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPricing.RowEnter

        Try
            If blnDeleting Then Exit Sub

            With dgvPricing.Rows(e.RowIndex)
                cmdDelete.Enabled = Not (.Cells(Columns.OE_Price_ID).Value Is Nothing)

                If .Cells(Columns.OE_Price_ID).Value.Equals(DBNull.Value) Then
                    If Not m_oPricing Is Nothing Then
                        If m_oPricing.Dirty Then
                            m_oPricing.Save()
                        End If
                    End If
                    m_oPricing = New cOEFlashPricing(m_oOrder.Ordhead.Cus_No) ' m_oOrder)
                Else
                    If Not m_oPricing Is Nothing Then
                        If m_oPricing.Dirty Then
                            m_oPricing.Save()
                        End If
                    End If
                    m_oPricing = New cOEFlashPricing(CInt(.Cells(Columns.OE_Price_ID).Value))

                End If

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvPricing_RowLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPricing.RowLeave

        If Not blnLoading Then
            Call Save()
        End If

    End Sub

    Private Sub SetColumnIndexSub(ByRef pdgv As DataGridView, ByVal columnIndex As Integer)

        pdgv.CurrentCell = pdgv.CurrentRow.Cells(columnIndex)

    End Sub

    Private Sub LineInsert(ByRef pdgv As DataGridView, ByRef pdt As DataTable, ByVal pintCell As Integer)

        Dim drNewRow As DataRow
        pdgv.AllowUserToAddRows = True

        drNewRow = pdt.NewRow

        pdt.Rows.Add(drNewRow)

        pdgv.AllowUserToAddRows = False
        pdgv.CurrentCell = pdgv.Rows(pdgv.Rows.Count - 1).Cells(pintCell) ' Columns.Item_No

        pdgv.Focus()

    End Sub

    Private Sub DeleteGridLine(Optional ByVal pblnAsk As Boolean = True)

        Try

            If dgvPricing.Rows.Count = 0 Then Exit Sub
            If dgvPricing.Rows(0).Cells(Columns.Program_Name).Value.ToString = "" Then Exit Sub

            ' ask if it is ok to delete. If not OK, then exit sub right away
            If pblnAsk Then
                If Trim(dgvPricing.CurrentRow.Cells(Columns.Program_Name).Value.ToString) <> "" Then
                    Dim mbrOkToDelete As MsgBoxResult
                    mbrOkToDelete = MsgBox("OK to delete ?", MsgBoxStyle.YesNo, "Order entry interface")

                    If mbrOkToDelete = MsgBoxResult.No Then Exit Sub
                End If
            End If

            Dim iPos As Integer = dgvPricing.CurrentRow.Index

            blnDeleting = True

            If dgvPricing.CurrentRow.Index >= 0 Then
                m_oPricing.Delete()
                m_oPricing = Nothing

                blnDeleting = False

                dgvPricing.Rows.RemoveAt(iPos)

            End If

            If dgvPricing.Rows.Count > 0 Then
                If iPos >= dgvPricing.Rows.Count Then
                    dgvPricing.CurrentCell = dgvPricing.Rows(0).Cells(Columns.Program_Name)
                    dgvPricing.CurrentCell = dgvPricing.Rows(iPos - 1).Cells(Columns.Program_Name)
                Else
                    If iPos < 0 Then iPos = 0
                    dgvPricing.CurrentCell = dgvPricing.Rows(iPos).Cells(Columns.Program_Name)
                End If
            End If

            dgvPricing.Focus()

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        blnDeleting = False

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

        Try
            If dgvPricing.Rows.Count = 0 Then Exit Sub

            dgvPricing.CurrentCell = dgvPricing.CurrentRow.Cells(Columns.Program_Name)

            Call DeleteGridLine()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub chkAllDates_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllDates.CheckedChanged

        Call LoadGrid()

    End Sub

    Private Sub cmdChargeAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChargeAdd.Click

        Try

            'blnInserting = True
            'Call ProductLineInsert()

            If dgvPricing.Rows.Count = 0 Then
                Call LineInsert(dgvPricing, m_dt, Columns.Program_Name)
            ElseIf Not (dgvPricing.Rows(dgvPricing.Rows.Count - 1).Cells(Columns.Program_Name).Value.Equals(DBNull.Value)) Then
                If Not Trim(dgvPricing.Rows(dgvPricing.Rows.Count - 1).Cells(Columns.Program_Name).Value).Equals("") Then
                    Call LineInsert(dgvPricing, m_dt, Columns.Program_Name)
                Else
                    dgvPricing.CurrentCell = dgvPricing.Rows(dgvPricing.Rows.Count - 1).Cells(Columns.Program_Name)
                End If
            Else
                dgvPricing.CurrentCell = dgvPricing.Rows(dgvPricing.Rows.Count - 1).Cells(Columns.Program_Name)
            End If

            'blnInserting = False

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Enum Columns

        OE_Price_ID
        Cus_No
        Program_Name
        Item_CD
        Prc_Or_Disc_1
        EQP_Type
        EQP_Plus_Pct
        Minimum_Qty_1
        Start_Date
        End_Date
        OEI_Warning
        Macola_Import
        When_Req
        Autorized_By
        Comments

    End Enum

End Class
