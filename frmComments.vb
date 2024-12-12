Public Class frmComments

    Private m_Ord_Guid As String
    'Private m_Ord_No As String
    'Private m_Cmt As String
    'Private m_Line_No As Integer
    'Private m_ID As Integer

    Private m_Comment As New cComment

    Public Sub New(ByVal pstrOrd_Guid As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        m_Ord_Guid = pstrOrd_Guid
        Call LoadGrid()

        cmdNew.Focus()

    End Sub

    Private Sub LoadGrid()

        dgvComments.DataSource = m_Comment.Load(m_Ord_Guid)
        txtOEI_Ord_No.Text = m_Comment.OEI_Ord_No ' dgvComments.Rows(0).Cells(0).Value.ToString

        dgvComments.Columns(Columns.ID).Visible = False
        dgvComments.Columns(Columns.Ord_Guid).Visible = False
        dgvComments.Columns(Columns.Line_Seq_No).Visible = False
        dgvComments.Columns(Columns.OEI_Ord_No).Visible = False
        dgvComments.Columns(Columns.Ord_No).Visible = False
        dgvComments.Columns(Columns.Cmt).HeaderText = "Comment"

        If dgvComments.Rows.Count <> 0 Then
            dgvComments.CurrentCell = dgvComments.Rows(0).Cells(Columns.Cmt)
            dgvComments.Focus()
        Else
            cmdNew.PerformClick()
        End If

        dgvComments.RowHeadersVisible = False
        dgvComments.Columns(Columns.Cmt).Width = dgvComments.Width - 4

    End Sub

    Public Property Ord_Guid() As String
        Get
            Ord_Guid = m_Ord_Guid
        End Get
        Set(ByVal value As String)
            m_Ord_Guid = value
        End Set
    End Property

    Private Sub dgvComments_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvComments.DoubleClick

        txtNew_Comment.Text = dgvComments.CurrentRow.Cells(Columns.Cmt).Value

    End Sub

    Private Sub txtNew_Comment_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNew_Comment.GotFocus

        If Trim(txtNew_Comment.Text) = "" And dgvComments.Rows.Count = 0 Then

            cmdNew.PerformClick()

        End If

    End Sub

    Private Sub txtNew_Comment_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNew_Comment.Validated

        If Trim(txtNew_Comment.Text) = "" Then Exit Sub

    End Sub

    Private Sub dgvComments_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvComments.RowEnter

        If dgvComments.Rows.Count = 0 Then Exit Sub

        If dgvComments.CurrentRow Is Nothing Then Exit Sub

        m_Comment = New cComment(CInt(dgvComments.Rows(e.RowIndex).Cells(Columns.ID).Value))

        txtNew_Comment.Text = m_Comment.Cmt

    End Sub

    Public Enum Columns

        ID
        Ord_Guid
        Cmt
        Line_Seq_No
        OEI_Ord_No
        Ord_No

    End Enum

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            If m_Comment.ID = 0 Then Exit Sub

            If Trim(txtNew_Comment.Text) = "" Then Exit Sub

            m_Comment.Cmt = txtNew_Comment.Text

            m_Comment.Save()
            Call LoadGrid()

            txtNew_Comment.Text = ""
            'txtNew_Comment.ReadOnly = True
            m_Comment.ID = 0

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            'txtNew_Comment.ReadOnly = False

            m_Comment = New cComment
            m_Comment.ID = -1
            m_Comment.Ord_Guid = m_Ord_Guid

            txtNew_Comment.Text = ""
            txtNew_Comment.Focus()

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            If m_Comment.ID <= 0 Then Exit Sub

            Dim oRep As New MsgBoxResult

            oRep = MsgBox("Do you want to delete this comment ?", MsgBoxStyle.YesNo, "Delete a comment")

            If oRep = MsgBoxResult.Yes Then

                m_Comment.Delete()
                Call LoadGrid()

            End If

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Me.Close()

    End Sub

    Private Sub txtNew_Comment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNew_Comment.KeyDown

        Try

            Select Case e.KeyValue

                Case Keys.Return

                    cmdSave.PerformClick()
                    cmdNew.PerformClick()

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub frmComments_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        dgvComments.Height = Me.Height - 171
        panButtons.Location = New Point(13, Me.Height - 74)

    End Sub

    Private Sub cmdAutoComments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAutoComments.Click

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            Dim oAuto As New frmCommentsAuto(m_oOrder.Ordhead.Ord_GUID, m_oOrder.Ordhead.OEI_Ord_No)

            oAuto.ShowDialog()

            Call LoadGrid()

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdCopyFromPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyFromPO.Click

        Dim strSql As String
        Dim db As New cDBA

        Try
            If Trim(txtCopyFromPO.Text) <> "" Then

                Dim dt As DataTable

                strSql = "" & _
                "SELECT     TOP 1 O.*     " & _
                "FROM       OELinCmt_Sql C " & _
                "INNER JOIN OEOrdHdr_Sql O ON C.Ord_No = O.Ord_No " & _
                "WHERE      O.OE_PO_No = '" & SqlCompliantString(Trim(txtCopyFromPO.Text)) & "' AND " & _
                "           O.Cus_No = '" & Trim(m_oOrder.Ordhead.Cus_No) & "' "

                dt = db.DataTable(strSql)
                If dt.Rows.Count <> 0 Then

                    strSql = _
                    "INSERT INTO OEI_LinCmt (Ord_Guid, Cmt) " & _
                    "SELECT     '" & m_oOrder.Ordhead.Ord_GUID & "', Cmt " & _
                    "FROM       OELinCmt_Sql " & _
                    "WHERE      Ord_No = '" & dt.Rows(0).Item("Ord_No").ToString & "' AND " & _
                    "           Ord_Type = '" & dt.Rows(0).Item("Ord_Type").ToString & "' " & _
                    "ORDER BY   Cmt_Seq_No "

                    db.Execute(strSql)

                Else

                    strSql = "" & _
                    "SELECT     C.*     " & _
                    "FROM       OEI_LinCmt C " & _
                    "INNER JOIN OEI_OrdHdr O ON C.Ord_Guid = O.Ord_Guid " & _
                    "WHERE      O.OE_PO_No = '" & SqlCompliantString(Trim(txtCopyFromPO.Text)) & "' AND " & _
                    "           O.Cus_No = '" & Trim(m_oOrder.Ordhead.Cus_No) & "' "

                    dt = db.DataTable(strSql)
                    If dt.Rows.Count <> 0 Then

                        strSql = _
                        "INSERT INTO OEI_LinCmt (Ord_Guid, Cmt) " & _
                        "SELECT '" & m_oOrder.Ordhead.Ord_GUID & "', Cmt " & _
                        "FROM OEI_LinCmt WHERE Ord_Guid = '" & Trim(dt.Rows(0).Item("Ord_Guid").ToString) & "' "

                        db.Execute(strSql)

                    End If

                End If

            End If

            strSql = _
            "UPDATE OEI_LinCmt " & _
            "SET Line_Seq_No = ID " & _
            "WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' AND Line_Seq_No IS NULL "

            db.Execute(strSql)

            Call LoadGrid()

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdMoveUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveUp.Click

        Dim intRow As Integer = dgvComments.CurrentRow.Index
        If dgvComments.Rows.Count > 1 Then

            If intRow > 0 Then

                Dim strSql1 As String = "UPDATE OEI_LinCmt SET Line_Seq_No = " & dgvComments.Rows(intRow - 1).Cells(Columns.Line_Seq_No).Value & " WHERE ID = " & dgvComments.Rows(intRow).Cells(Columns.ID).Value
                Dim strSql2 As String = "UPDATE OEI_LinCmt SET Line_Seq_No = " & dgvComments.Rows(intRow).Cells(Columns.Line_Seq_No).Value & " WHERE ID = " & dgvComments.Rows(intRow - 1).Cells(Columns.ID).Value

                Dim db As New cDBA

                'Debug.Print(strSql1)
                db.Execute(strSql1)
                'Debug.Print(strSql1)
                db.Execute(strSql2)

                dgvComments.DataSource = m_Comment.Load(m_Ord_Guid)

                dgvComments.CurrentCell = dgvComments.Rows(intRow - 1).Cells(Columns.Cmt)

            End If

        End If

    End Sub


    Private Sub cmdMoveDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveDown.Click

        Dim intRow As Integer = dgvComments.CurrentRow.Index
        If dgvComments.Rows.Count > 1 Then

            If intRow < (dgvComments.Rows.Count - 1) Then

                Dim strSql1 As String = "UPDATE OEI_LinCmt SET Line_Seq_No = " & dgvComments.Rows(intRow + 1).Cells(Columns.Line_Seq_No).Value & " WHERE ID = " & dgvComments.Rows(intRow).Cells(Columns.ID).Value
                Dim strSql2 As String = "UPDATE OEI_LinCmt SET Line_Seq_No = " & dgvComments.Rows(intRow).Cells(Columns.Line_Seq_No).Value & " WHERE ID = " & dgvComments.Rows(intRow + 1).Cells(Columns.ID).Value

                Dim db As New cDBA

                'Debug.Print(strSql1)
                db.Execute(strSql1)
                'Debug.Print(strSql1)
                db.Execute(strSql2)

                dgvComments.DataSource = m_Comment.Load(m_Ord_Guid)

                dgvComments.CurrentCell = dgvComments.Rows(intRow + 1).Cells(Columns.Cmt)

            End If

        End If

    End Sub

    Private Sub cmdDeleteAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteAll.Click

        Try

            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            'If m_Comment.ID <= 0 Then Exit Sub

            Dim oRep As New MsgBoxResult

            oRep = MsgBox("Do you want to delete every comments ?", MsgBoxStyle.YesNo, "Delete a comment")

            If oRep = MsgBoxResult.Yes Then

                Dim db As New cDBA
                Dim strSql As String = _
                "DELETE FROM OEI_LinCmt WHERE Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' "

                db.Execute(strSql)

                Call LoadGrid()

            End If

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

End Class