Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.VisualBasic.Compatibility.VB6

Friend Class frmOEFlash
    Inherits System.Windows.Forms.Form

    Private dtLocal As DataTable
    Private dtCustomerInfo As DataTable
    Private dtComments As DataTable

    Private m_Cus_No As String

    Private db As New cDBA

    'Private daDataAdapter As OleDb.OleDbDataAdapter
    'Private dtComments As DataTable

    'Constants for topmost.
    Private Const HWND_TOPMOST As Short = -1
    Private Const HWND_NOTOPMOST As Short = -2
    Private Const SWP_NOMOVE As Integer = &H2
    Private Const SWP_NOSIZE As Integer = &H1
    Private Const SWP_NOACTIVATE As Integer = &H10
    Private Const SWP_SHOWWINDOW As Integer = &H40
    Private Const FLAGS As Boolean = SWP_NOMOVE Or SWP_NOSIZE

    Const MOUSEEVENTF_LEFTDOWN As Short = 2
    Const MOUSEEVENTF_LEFTUP As Short = 4
    Const MOUSEEVENTF_RIGHTDOWN As Short = 8
    Const MOUSEEVENTF_RIGHTUP As Short = 16

    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cbuttons As Integer, ByVal dwExtraInfo As Integer)

    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

    Public Sub New(ByVal pstrCus_No As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Cus_No = pstrCus_No

    End Sub

    Public Property Cus_No() As String
        Get
            Cus_No = m_Cus_No
        End Get
        Set(ByVal value As String)
            m_Cus_No = value
            txtCusNo.Text = m_Cus_No
            Call LoadCustomerInfo()
            Call LoadGrdComments()
            Call SetGrdCommentsColumns()
        End Set
    End Property

    ' FORM LOAD
    Private Sub frmOEFlash_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Try

            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'Display Version
            'lblVersion.Text = gVersion

            ' Disable Comment Editing buttons by default
            cmdInsertComment.Enabled = False
            cmdDeleteComment.Enabled = False

            Call LoadCustomerInfo()
            Call LoadGrdComments()
            Call SetGrdCommentsColumns()

            Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

            'mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' FORM UNLOAD
    Private Sub frmOEFlash_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Call SAVEFormPos()

    End Sub

    ' FORM ACTIVATE
    Private Sub frmOEFlash_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        Call GETFormPos()

    End Sub

    ' FORM DE-ACTIVATE
    Private Sub frmOEFlash_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate

        Call SAVEFormPos()

    End Sub

    ' FORM QUERY UNLOAD
    Private Sub frmOEFlash_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        '    If UnloadMode = vbFormControlMenu Then
        '        MsgBox "Please use buttons to close form", vbOKOnly, "Can't close form"
        '        Cancel = True
        '    End If
        eventArgs.Cancel = Cancel
    End Sub


    Private Sub cmdRefresh_Click()

        Try
            Call LoadCustomerInfo()
            Call LoadGrdComments()
            Call SetGrdCommentsColumns()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub txtCusNo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCusNo.Leave

        Try
            Call LoadCustomerInfo()
            Call LoadGrdComments()
            Call SetGrdCommentsColumns()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub chkUnlock_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkUnlock.CheckStateChanged

        Try
            If txtCusNo.Text = "" Or txtCusName.Text = "" Then
                MsgBox("You must first select a customer!", MsgBoxStyle.OkOnly, "Select Customer")
                Exit Sub
            End If

            If chkUnlock.CheckState = 1 Then
                chkUnlock.Text = "Click to Lock from Editing"
                'dgvComments.Columns(1).ReadOnly = False
                fraComments.Text = "Customer Comments (OPEN FOR EDITING)"
            Else
                chkUnlock.Text = "Click to Unlock for Editing"
                'dgvComments.Columns(1).ReadOnly = True
                fraComments.Text = "Customer Comments"
            End If

            cmdInsertComment.Enabled = chkUnlock.CheckState
            cmdDeleteComment.Enabled = chkUnlock.CheckState
            cmdUp.Enabled = chkUnlock.CheckState
            cmdDown.Enabled = chkUnlock.CheckState

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub



    Private Sub cmdInsertComment_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInsertComment.Click

        Dim dt As DataTable

        Dim sql As String
        Dim cmtID As String
        Dim sortID As String = ""
        ' Dim prevCmtID As String
        Dim newID As String
        Dim bEmpty As Boolean
        Dim rowInsert As Short
        bEmpty = False

        Try

            Call UpdateLines()

            dgvComments.AllowUserToAddRows = True




            '--------------it was until 5.30.2018 but I put in comments because = nothing is not correct -------------------
            ' START Find the correct row to insert after
            'If dgvComments.Rows.Count <> 0 AndAlso dgvComments.CurrentRow.Index = Nothing Then
            '---------------------------------------------------
            If dgvComments.Rows.Count <> 0 AndAlso dgvComments.CurrentRow Is Nothing Then



                'sql = "SELECT ID, Cus_Sort FROM dbo.Exact_Custom_OeFlash_CustomerComments WITH (Nolock) WHERE Cus_no = '" & txtCusNo.Text & "' " & " ORDER BY Cus_Sort"
                'rsLocal = New ADODB.Recordset
                ''    rsLocal.CursorLocation = adUseClient   ' NEED THIS to sort.  If not sorting, ignore this line.
                'rsLocal.Open(sql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                ' END Find row

                ' if customer already has comments
                'If Not (rsLocal.BOF And rsLocal.EOF) Then

                '++ID 5.30.2018 modified
                rowInsert = dgvComments.Rows(dgvComments.Rows.Count - 1).Index  '.CurrentRow.Index

                ' rowInsert = dgvComments.CurrentRow.Index
                ' If rowInsert < 0 Then rowInsert = 0
                'rsLocal.Move((rowInsert))

                If Not chkInsertAtEnd.Checked Then
                    ''MsgBox rowInsert
                    'cmtID = dgvComments.Rows(dgvComments.CurrentRow.Index).Cells(0).Value ' rsLocal.Fields("ID").Value ' Retrieve ID of comment to delete
                    'sortID = dgvComments.Rows(dgvComments.CurrentRow.Index).Cells(2).Value - 1 ' rsLocal.Fields("Cus_Sort").Value

                    cmtID = dgvComments.Rows(rowInsert).Cells(0).Value ' rsLocal.Fields("ID").Value ' Retrieve ID of comment to delete
                    sortID = dgvComments.Rows(rowInsert).Cells(2).Value - 1 ' rsLocal.Fields("Cus_Sort").Value



                    ' START Update Cus_Sort for next row
                    sql = "UPDATE dbo.Exact_Custom_OeFlash_CustomerComments " & " SET Cus_Sort = Cus_Sort + 2 " & " WHERE Cus_Sort > '" & sortID & "'"
                    db.Execute(sql)
                    'rsLocal = New ADODB.Recordset
                    'rsLocal.Open(sql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                    'rsLocal.Close()
                End If
                ' END Update row
            Else
                bEmpty = True
            End If

            ' START Insert row
            '   find last created ID in comments table
            If chkInsertAtEnd.Checked Then
                sql = "SELECT MAX(Cus_Sort) AS ID FROM dbo.Exact_Custom_OeFlash_CustomerComments WITH (Nolock) "
            Else
                sql = "SELECT MAX(ID) AS ID FROM dbo.Exact_Custom_OeFlash_CustomerComments WITH (Nolock) "
            End If

            newID = 0
            dt = db.DataTable(sql)
            If dt.Rows.Count <> 0 Then
                newID = dt.Rows(0).Item("ID") + 1
            End If

            'rsLocal = New ADODB.Recordset
            'rsLocal.Open(sql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            'newID = rsLocal.Fields("ID").Value + 1 ' Retrieve ID of comment to delete

            ' if no comments exist, sortID will be the same as commentID
            If bEmpty Or chkInsertAtEnd.Checked Then
                sortID = newID
            End If

            '   Allow Identity inserts on table
            '    sql = "SET IDENTITY_INSERT dbo.Exact_Custom_OeFlash_CustomerComments ON"
            '    Set rsLocal = New ADODB.Recordset
            '    rsLocal.Open sql, gConn.ConnectionString, adOpenDynamic, adLockOptimistic

            '   Insert new comment line into table
            sql = "INSERT INTO dbo.Exact_Custom_OeFlash_CustomerComments ( Cus_no, Cmt, Cus_Sort ) " & " VALUES ( '" & SqlCompliantString(txtCusNo.Text) & "', '', '" & (CDbl(sortID) + 1) & "' )"
            db.Execute(sql)
            'rsLocal = New ADODB.Recordset
            'rsLocal.Open(sql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            ' END Insert row

            dgvComments.AllowUserToAddRows = False

            ' Update Grid
            Call LoadGrdComments()
            Call SetGrdCommentsColumns()

            If chkInsertAtEnd.Checked Then
                rowInsert = dgvComments.Rows.Count - 1
            End If
            'If rowInsert + 1 > dgvComments.Rows.Count Then rowInsert = rowInsert - 1
            dgvComments.CurrentCell = dgvComments(1, CInt(dgvComments.Rows.Count - 1))

            dgvComments.BeginEdit(True)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub



    Private Sub cmdDeleteComment_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDeleteComment.Click

        'Dim sql As String
        'Dim cmtID As String
        'Dim lineNo As Short
        Dim nbRows As Short

        Try

            Call UpdateLines()

            nbRows = dtComments.Rows.Count

            If nbRows = 0 Then Exit Sub

            Dim sqlCommand As New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
            db.DBDataAdapter.DeleteCommand = sqlCommand.GetDeleteCommand
            db.DBDataTable.Rows(dgvComments.CurrentRow.Index).Delete() ' .RemoveAt(0)
            db.DBDataAdapter.Update(db.DBDataTable)

            '' START Delete row
            'sql = "DELETE FROM Exact_Custom_OeFlash_CustomerComments WHERE ID = '" & Trim(cmtID) & "'"

            '' START Find the correct row to delete
            'sql = "SELECT ID " & " FROM Exact_Custom_OeFlash_CustomerComments " & " WHERE Cus_no = '" & txtCusNo.Text & "' " & " ORDER BY Cus_Sort"
            'rsLocal = New ADODB.Recordset
            'rsLocal.CursorLocation = ADODB.CursorLocationEnum.adUseClient ' NEED THIS to sort.  If not sorting, ignore this line.
            'rsLocal.Open(sql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

            '' Make delete only if records exist
            'If Not (rsLocal.BOF And rsLocal.EOF) Then

            '    lineNo = dgvComments.CurrentRow.Index - 1
            '    rsLocal.Move((lineNo))
            'cmtID = rsLocal.Fields("ID").Value ' Retrieve ID of comment to delete

            '' START Delete row
            'sql = "DELETE FROM Exact_Custom_OeFlash_CustomerComments WHERE ID = '" & Trim(cmtID) & "'"
            'rsLocal = New ADODB.Recordset
            'rsLocal.Open(sql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            '' END Delete row

            '' Update Grid
            'Call LoadGrdComments()
            'Call SetGrdCommentsColumns()

            'If lineNo = nbRows - 1 Then
            '    dgvComments.CurrentCell = dgvComments(0, lineNo - 1)
            'Else
            '    dgvComments.CurrentCell = dgvComments(0, lineNo)
            'End If

            'End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub cmdCloseScreen_Click()

        Me.Close()

    End Sub


    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click

        Try
            Call UpdateLines()

            Me.Close()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub LoadGrdComments()

        Dim sql As String

        Try
            sql = "SELECT ID, Cmt AS Comments, Cus_Sort FROM Exact_Custom_OeFlash_CustomerComments WHERE Cus_no = '" & txtCusNo.Text & "' ORDER BY Cus_Sort"

            'daDataAdapter = New OleDb.OleDbDataAdapter
            dtComments = New DataTable

            db = New cDBA()

            'daDataAdapter = dbConn.DataAdapter(sql)
            dtComments = db.DataTable(sql)
            dgvComments.DataSource = dtComments

            'rsComments = New ADODB.Recordset
            ' ''rsComments.CursorLocation = ADODB.CursorLocationEnum.adUseClient ' NEED THIS to sort.  If not sorting, ignore this line.
            ' ''            rsComments.Open(sql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            'rsComments.Open(sql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            ''datComments = Nothing
            'datComments.Recordset = rsComments
            ' ''dgvComments.DataSource = rsComments
            ' ''dgvComments.Refresh()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SetGrdCommentsColumns()

        Try

            dgvComments.Columns(0).Visible = False ' 0
            dgvComments.Columns(1).Width = dgvComments.Width
            dgvComments.Columns(2).Visible = False ' .Width = 0
            'dgvComments.Columns.Add(DGVTextBoxColumn("Cmd", "Comments", 200))

            'With grdComments
            '    With .Columns(0)
            '        .DataField = "Comment"
            '        .Caption = "Comment"
            '        .Width = .TwipsToPixelsX(10500)
            '        .Alignment = MSDataGridLib.AlignmentConstants.dbgLeft
            '        If chkUnlock.CheckState = 1 Then
            '            .Locked = False
            '        Else
            '            .Locked = True
            '        End If
            '    End With
            'End With

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadCustomerInfo()

        Dim sql As String

        Try
            'lblVersion.Text = gVersion

            'Modified July 22, 2006 - Add CSR
            sql = _
            "SELECT cmp_name, ISNULL(ClassificationID, '') AS ClassificationID, ISNULL(TextField1, '') AS EQP  " & _
            "FROM   cicmpy WITH (Nolock) " & _
            "WHERE  cmp_code = '" & txtCusNo.Text & "'"

            dtCustomerInfo = db.DataTable(sql)

            'rsCustomerInfo = New ADODB.Recordset
            'rsCustomerInfo.Open(sql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

            If dtCustomerInfo.Rows.Count <> 0 Then ' Not (rsCustomerInfo.BOF And rsCustomerInfo.EOF) Then
                txtCusName.Text = Trim(dtCustomerInfo.Rows(0).Item("cmp_name"))
                txtClassID.Text = "Star Lvl: " & dtCustomerInfo.Rows(0).Item("ClassificationID")
                If Trim(dtCustomerInfo.Rows(0).Item("EQP")) = "" Then
                    txtEQP.Text = ""
                Else
                    txtEQP.Text = "Terms: " & dtCustomerInfo.Rows(0).Item("EQP")
                End If
                chkUnlock.Enabled = True
                Call chkUnlock_CheckStateChanged(chkUnlock, New System.EventArgs())
            Else
                txtCusName.Text = ""
                txtClassID.Text = ""
                txtEQP.Text = ""
                chkUnlock.Enabled = False
            End If

            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'rsCustomerInfo.Close()
            ''UPGRADE_NOTE: Object rsCustomerInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            'rsCustomerInfo = Nothing

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub UpdateLines()

        Try

            Dim sqlCommand As New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
            db.DBDataAdapter.UpdateCommand = sqlCommand.GetUpdateCommand
            db.DBDataAdapter.Update(db.DBDataTable)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////       SAVE / GET FORM POSITION             ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub GETFormPos()

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            '    With frmXXXXXXXXXXXPos
            '        If .Modified = True Then
            '            Me.Left = .Left
            '            Me.Top = .Top
            '        End If
            '    End With
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SAVEFormPos()

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            '    With frmXXXXXXXXXXPos     ' Add variable to modGlobal
            '        .Left = Me.Left
            '        .Top = Me.Top
            '        .Modified = True
            '    End With

            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub dgvComments_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvComments.CellBeginEdit

        ' if checkbox unlock
        If chkUnlock.CheckState = 0 Then e.Cancel = True

        ' We don't want to overwrite the ID field
        If e.ColumnIndex = 0 Then e.Cancel = True

    End Sub

    Private Sub frmOEFlash_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        fraComments.Height = Me.Height - 147
        dgvComments.Height = Me.Height - 170

    End Sub

    Private Sub cmdUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUp.Click

        Call MoveGridElementUp()

    End Sub

    Private Sub cmdDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDown.Click

        Call MoveGridElementDown()

    End Sub

    Private Sub MoveGridElementDown()

        Dim strSql As String
        Dim intUpped As Integer
        Dim intDowned As Integer
        Dim intDownedPosition As Integer

        Try
            ' Do nothing on the last row
            If dgvComments.Rows.Count < 2 Then Exit Sub
            If dgvComments.CurrentRow.Index = (dgvComments.Rows.Count - 1) Then Exit Sub

            Call UpdateLines()

            intDowned = dgvComments.Rows(dgvComments.CurrentRow.Index + 1).Cells(2).Value
            intUpped = dgvComments.Rows(dgvComments.CurrentRow.Index).Cells(2).Value
            intDownedPosition = dgvComments.CurrentRow.Index + 1

            strSql = "UPDATE dbo.Exact_Custom_OeFlash_CustomerComments " & " SET Cus_Sort = " & intDowned & " WHERE ID = " & dgvComments.Rows(dgvComments.CurrentRow.Index).Cells(0).Value
            db.Execute(strSql)
            'rsLocal = New ADODB.Recordset
            'rsLocal.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            strSql = "UPDATE dbo.Exact_Custom_OeFlash_CustomerComments " & " SET Cus_Sort = " & intUpped & " WHERE ID = " & dgvComments.Rows(dgvComments.CurrentRow.Index + 1).Cells(0).Value
            db.Execute(strSql)
            'rsLocal = New ADODB.Recordset
            'rsLocal.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            Call LoadGrdComments()

            dgvComments.CurrentCell = dgvComments(1, intDownedPosition)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub MoveGridElementUp()

        Dim strSql As String
        Dim intUpped As Integer
        Dim intDowned As Integer
        Dim intUppedPosition As Integer

        Try
            ' Do nothing on the first row
            If dgvComments.Rows.Count < 2 Then Exit Sub
            If dgvComments.CurrentRow.Index = 0 Then Exit Sub

            Call UpdateLines()

            intUpped = dgvComments.Rows(dgvComments.CurrentRow.Index - 1).Cells(2).Value
            intDowned = dgvComments.Rows(dgvComments.CurrentRow.Index).Cells(2).Value
            intUppedPosition = dgvComments.CurrentRow.Index - 1

            strSql = "UPDATE dbo.Exact_Custom_OeFlash_CustomerComments " & " SET Cus_Sort = " & intUpped & " WHERE ID = " & dgvComments.Rows(dgvComments.CurrentRow.Index).Cells(0).Value
            db.Execute(strSql)
            'rsLocal = New ADODB.Recordset
            'rsLocal.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            strSql = "UPDATE dbo.Exact_Custom_OeFlash_CustomerComments " & " SET Cus_Sort = " & intDowned & " WHERE ID = " & dgvComments.Rows(dgvComments.CurrentRow.Index - 1).Cells(0).Value
            db.Execute(strSql)
            'rsLocal = New ADODB.Recordset
            'rsLocal.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            Call LoadGrdComments()

            dgvComments.CurrentCell = dgvComments(1, intUppedPosition)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

End Class