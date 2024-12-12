Public Class ucOEFlashCharges

    Private m_strCus_No As String = String.Empty
    Private db As New cDBA

    Private m_oCharges As cOEFlashCharges
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
                If Not m_oCharges Is Nothing Then

                    If m_oCharges.Dirty Then
                        m_oCharges.Save()
                    End If

                    With dgvCharges.CurrentRow.Cells(Columns.Charge_Usage_ID)
                        If .Value.Equals(DBNull.Value) Then
                            .Value = m_oCharges.Charge_Usage_ID
                        ElseIf .Value <> m_oCharges.Charge_Usage_ID Then
                            .Value = m_oCharges.Charge_Usage_ID
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
            If dgvCharges.Columns.Count <> 0 Then Exit Sub

            With dgvCharges.Columns

                .Add(DGVTextBoxColumn("Charge_Usage_ID", "Charge_Usage_ID", 0))
                .Add(DGVTextBoxColumn("Charge_CD", "Charge", 115))
                .Add(DGVTextBoxColumn("Cus_No", "Cus_No", 0))
                .Add(DGVCheckBoxColumn("Always_Use", "Always", 42))
                .Add(DGVCheckBoxColumn("Never_Use", "Never", 35))
                .Add(DGVCheckBoxColumn("No_Charge", "N/C", 25))
                .Add(DGVCalendarColumn("Charge_From", "From", 70))
                .Add(DGVCalendarColumn("Charge_To", "To", 70))
                .Add(DGVCheckBoxColumn("Blind", "Blind", 27))
                .Add(DGVTextBoxColumn("Apply_To_Ship_To", "On ShipTo", 50))
                .Add(DGVTextBoxColumn("Apply_To_Program", "On Program", 50))
                .Add(DGVTextBoxColumn("When_Qty_LT", "Qty GT", 40))
                .Add(DGVTextBoxColumn("When_Qty_GT", "Qty LT", 40))
                .Add(DGVTextBoxColumn("When_Amt_LT", "Amt GT", 40))
                .Add(DGVTextBoxColumn("When_Amt_GT", "Amt LT", 40))
                .Add(DGVCheckBoxColumn("When_Req", "When Req", 32))
                .Add(DGVTextBoxColumn("Autorized_By", "Autorized By", 75))
                .Add(DGVTextBoxColumn("Comments", "Comments", 200))

            End With

            dgvCharges.Columns(Columns.Charge_Usage_ID).Visible = False
            dgvCharges.Columns(Columns.Cus_No).Visible = False
            dgvCharges.Columns(Columns.Apply_To_Ship_To).Visible = False
            dgvCharges.Columns(Columns.Apply_To_Program).Visible = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadGrid()

        Dim strSql As String

        Try
            strSql = _
            "SELECT		CU.CHARGE_USAGE_ID, CU.Charge_CD, CU.Cus_No, CU.Always_Use, " & _
            "           CU.Never_Use, CU.No_Charge, ISNULL(CU.Charge_From, GETDATE()) AS Charge_From, " & _
            "           ISNULL(CU.Charge_To, GETDATE()) AS Charge_To, CU.Blind, CU.Apply_To_Ship_To, CU.Apply_To_Program, " & _
            "           CU.When_Qty_GT, CU.When_Qty_LT, CU.When_Amt_GT, CU.When_Amt_LT, " & _
            "			CU.When_Req, CU.Autorized_By, CU.Comments " & _
            "FROM		MDB_CFG_CHARGE_USAGE CU WITH (Nolock) " & _
            "INNER JOIN	MDB_CFG_CHARGE C WITH (Nolock) ON CU.Charge_CD = C.Charge_CD " & _
            "WHERE		CU.Cus_No = '" & m_strCus_No & "' "
            If Not chkAllDates.Checked Then strSql &= " AND " & _
            "           ( " & _
            "				(CU.Charge_From IS NULL AND CU.Charge_To IS NULL) OR " & _
            "				(CU.Charge_From <= GETDATE()) AND CU.Charge_To >= DATEDIFF(d, 0, GETDATE()) OR " & _
            "				(CU.Charge_From <= GETDATE()) AND (CU.Charge_To IS NULL) OR " & _
            "				(CU.Charge_From IS NULL AND CU.Charge_To >= DATEDIFF(d, 0, GETDATE()))" & _
            "			) "
            strSql &= " ORDER BY CU.Charge_CD, CU.Charge_To DESC, CU.Charge_From "

            blnLoading = True

            m_dt = db.DataTable(strSql)
            dgvCharges.DataSource = m_dt

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

    Private Sub dgvCharges_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCharges.CellEnter

        Try

            If blnDeleting Then Exit Sub

            If dgvCharges.Rows.Count = 0 Then Exit Sub
            'If dgvItems.CurrentRow.Index = 0 Then Exit Sub

            If e.ColumnIndex < Columns.Charge_CD Then

                Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
                dgvCharges.BeginInvoke(oInvoke, dgvCharges, Columns.Charge_CD)

            ElseIf e.ColumnIndex = Columns.Comments Then ' LastCol

                Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
                dgvCharges.BeginInvoke(oInvoke, dgvCharges, Columns.Charge_CD)

                'ElseIf e.ColumnIndex > Columns.Item_No And IIf(dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Equals(DBNull.Value), "", dgvItems.Rows(e.RowIndex).Cells(Columns.Item_No).Value) = "" Then
            ElseIf e.ColumnIndex > Columns.Charge_CD And dgvCharges.Rows(e.RowIndex).Cells(Columns.Charge_CD).Value Is Nothing Then
                'ElseIf e.ColumnIndex > Columns.Item_No And IIf(dgvItems.CurrentRow.Cells(Columns.Item_No).Equals(DBNull.Value), "", dgvItems.CurrentRow.Cells(Columns.Item_No).Value) = "" Then

                Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
                dgvCharges.BeginInvoke(oInvoke, dgvCharges, Columns.Charge_CD)

            ElseIf e.ColumnIndex > Columns.Charge_CD And Trim(dgvCharges.Rows(e.RowIndex).Cells(Columns.Charge_CD).Value.ToString) = "" Then

                Dim oInvoke As New SetColumnIndex(AddressOf SetColumnIndexSub)
                dgvCharges.BeginInvoke(oInvoke, dgvCharges, Columns.Charge_CD)

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvCharges_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvCharges.CellValidating

        If blnLoading Then Exit Sub

        If m_oCharges Is Nothing Then
            m_oCharges = New cOEFlashCharges(m_oOrder.Ordhead.Cus_No) ' THAT CODE SHOULD NEVER HAPPEN
        End If

        Try
            Select Case e.ColumnIndex

                Case Columns.Charge_CD
                    If Trim(e.FormattedValue).ToString <> "" Then
                        If m_oCharges.Charge_Exist(Trim(e.FormattedValue).ToString) Then
                            m_oCharges.Charge_CD = Trim(e.FormattedValue).ToString
                        Else
                            Throw New OEException(OEError.Charge_Not_Found, True)
                        End If
                    End If

                    'Case Columns.Cus_No
                    '    m_oCharges.Cus_No = Trim(e.FormattedValue).ToString

                Case Columns.Always_Use
                    m_oCharges.Always_Use = (e.FormattedValue)

                Case Columns.Never_Use
                    m_oCharges.Never_Use = (e.FormattedValue)

                Case Columns.No_Charge
                    m_oCharges.No_Charge = (e.FormattedValue)

                Case Columns.Charge_From
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format, True)
                        m_oCharges.Charge_From = CDate(e.FormattedValue)
                    End If

                Case Columns.Charge_To
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsDate(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Date_Format, True)
                        m_oCharges.Charge_To = CDate(e.FormattedValue)
                    End If

                Case Columns.Blind
                    m_oCharges.Blind = (e.FormattedValue)

                Case Columns.Apply_To_Ship_To
                    m_oCharges.Apply_To_Ship_To = Trim(e.FormattedValue).ToString

                Case Columns.Apply_To_Program
                    m_oCharges.Apply_To_Program = Trim(e.FormattedValue).ToString

                Case Columns.When_Qty_GT
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                        m_oCharges.When_Qty_GT = (e.FormattedValue)
                    End If

                Case Columns.When_Qty_LT
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                        m_oCharges.When_Qty_LT = (e.FormattedValue)
                    End If

                Case Columns.When_Amt_GT
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                        m_oCharges.When_Amt_GT = (e.FormattedValue)
                    End If

                Case Columns.When_Amt_LT
                    If Trim(e.FormattedValue) <> "" Then
                        If Not IsNumeric(e.FormattedValue) Then Throw New OEException(OEError.Invalid_Number_Format, True)
                        m_oCharges.When_Amt_LT = (e.FormattedValue)
                    End If

                Case Columns.When_Req
                    m_oCharges.When_Req = (e.FormattedValue)

                Case Columns.Autorized_By
                    m_oCharges.Autorized_By = Trim(e.FormattedValue).ToString

                Case Columns.Comments
                    m_oCharges.Comments = Trim(e.FormattedValue).ToString

            End Select

        Catch oe_er As OEException
            e.Cancel = oe_er.Cancel
            If oe_er.ShowMessage Then MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try


    End Sub

    Private Sub dgvCharges_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvCharges.KeyDown

        Try
            Select Case dgvCharges.CurrentCell.ColumnIndex

                Case Columns.Charge_CD

                    If e.KeyCode = Keys.F7 Then

                        Dim oSearch As New cSearch("MDB-CFG-CHARGE")

                        If oSearch.Form.FoundRow <> -1 Then
                            dgvCharges.CurrentCell.Value = oSearch.Form.FoundElementValue

                        End If
                        oSearch = Nothing

                    End If

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvCharges_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCharges.LostFocus

        Try
            If m_oCharges Is Nothing Then Exit Sub

            If m_oCharges.Dirty Then
                m_oCharges.Save()
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvCharges_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCharges.RowEnter

        Try
            If blnDeleting Then Exit Sub

            With dgvCharges.Rows(e.RowIndex)
                cmdDelete.Enabled = Not (.Cells(Columns.Charge_Usage_ID).Value Is Nothing)

                If .Cells(Columns.Charge_Usage_ID).Value.Equals(DBNull.Value) Then
                    If Not m_oCharges Is Nothing Then
                        If m_oCharges.Dirty Then
                            m_oCharges.Save()
                        End If
                    End If
                    m_oCharges = New cOEFlashCharges(m_oOrder.Ordhead.Cus_No) ' m_oOrder)
                Else
                    If Not m_oCharges Is Nothing Then
                        If m_oCharges.Dirty Then
                            m_oCharges.Save()
                        End If
                    End If
                    m_oCharges = New cOEFlashCharges(CInt(.Cells(Columns.Charge_Usage_ID).Value))
                End If

            End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub dgvCharges_RowLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCharges.RowLeave

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

            If dgvCharges.Rows.Count = 0 Then Exit Sub
            If dgvCharges.Rows(0).Cells(Columns.Charge_CD).Value.ToString = "" Then Exit Sub

            ' ask if it is ok to delete. If not OK, then exit sub right away
            If pblnAsk Then
                If Trim(dgvCharges.CurrentRow.Cells(Columns.Charge_CD).Value.ToString) <> "" Then
                    Dim mbrOkToDelete As MsgBoxResult
                    mbrOkToDelete = MsgBox("OK to delete ?", MsgBoxStyle.YesNo, "Order entry interface")

                    If mbrOkToDelete = MsgBoxResult.No Then Exit Sub
                End If
            End If

            Dim iPos As Integer = dgvCharges.CurrentRow.Index

            blnDeleting = True

            If dgvCharges.CurrentRow.Index >= 0 Then
                m_oCharges.Delete()
                m_oCharges = Nothing

                blnDeleting = False

                dgvCharges.Rows.RemoveAt(iPos)

            End If

            If dgvCharges.Rows.Count > 0 Then
                If iPos >= dgvCharges.Rows.Count Then
                    dgvCharges.CurrentCell = dgvCharges.Rows(0).Cells(Columns.Charge_CD)
                    dgvCharges.CurrentCell = dgvCharges.Rows(iPos - 1).Cells(Columns.Charge_CD)
                Else
                    If iPos < 0 Then iPos = 0
                    dgvCharges.CurrentCell = dgvCharges.Rows(iPos).Cells(Columns.Charge_CD)
                End If
            End If

            dgvCharges.Focus()

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        blnDeleting = False

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

        Try
            If dgvCharges.Rows.Count = 0 Then Exit Sub

            dgvCharges.CurrentCell = dgvCharges.CurrentRow.Cells(Columns.Charge_CD)

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

            If dgvCharges.Rows.Count = 0 Then
                Call LineInsert(dgvCharges, m_dt, Columns.Charge_CD)
            ElseIf Not (dgvCharges.Rows(dgvCharges.Rows.Count - 1).Cells(Columns.Charge_CD).Value.Equals(DBNull.Value)) Then
                If Not Trim(dgvCharges.Rows(dgvCharges.Rows.Count - 1).Cells(Columns.Charge_CD).Value).Equals("") Then
                    Call LineInsert(dgvCharges, m_dt, Columns.Charge_CD)
                Else
                    dgvCharges.CurrentCell = dgvCharges.Rows(dgvCharges.Rows.Count - 1).Cells(Columns.Charge_CD)
                End If
            Else
                dgvCharges.CurrentCell = dgvCharges.Rows(dgvCharges.Rows.Count - 1).Cells(Columns.Charge_CD)
            End If

            'blnInserting = False

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Enum Columns
        Charge_Usage_ID
        Charge_CD
        Cus_No
        Always_Use
        Never_Use
        No_Charge
        Charge_From
        Charge_To
        Blind
        Apply_To_Ship_To
        Apply_To_Program
        When_Qty_GT
        When_Qty_LT
        When_Amt_GT
        When_Amt_LT
        When_Req
        Autorized_By
        Comments
    End Enum

End Class
