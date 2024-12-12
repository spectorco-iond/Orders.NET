Imports System.Data
Imports System.Collections.Generic
Public Class frm_route_maintain
    Dim dt As New DataTable

    Dim olist_US_Exact_Traveler_Route_Accepted As List(Of cUS_Exact_Traveler_Route_Accepted)
    Dim _cUS_Exact_Traveler_Route_Accepted_DAL As cUS_Exact_Traveler_Route_Accepted_DAL = New cUS_Exact_Traveler_Route_Accepted_DAL()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dgvshow()
    End Sub

    Private Sub dgvshow()
        olist_US_Exact_Traveler_Route_Accepted = _cUS_Exact_Traveler_Route_Accepted_DAL.GetAllList()
        DataGridView1.MultiSelect = False
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        initial_header_dgv()
        show_row_in_dgv()
        ButtonDEL.Visible = False
    End Sub

    Private Sub initial_header_dgv()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("US_Route_Accepted_ID", "US_Route_Accepted_ID")
        DataGridView1.Columns.Add("CAT", "CAT")
        DataGridView1.Columns.Add("RouteCatID", "RouteCatID")
        DataGridView1.Columns.Add("RouteCategory", "RouteCategory")
        DataGridView1.Columns.Add("deco_meth_1", "deco_meth_1")
        DataGridView1.Columns.Add("deco_meth_2", "deco_meth_2")
        DataGridView1.Columns.Add("RouteID", "RouteID")
        DataGridView1.Columns.Add("RouteDescription", "RouteDescription")

        DataGridView1.Columns("US_Route_Accepted_ID").Visible = False
        DataGridView1.Columns("US_Route_Accepted_ID").ReadOnly = True
        DataGridView1.Columns("RouteCatID").Visible = False
        DataGridView1.Columns("RouteCategory").Visible = False
        DataGridView1.Columns("deco_meth_1").Visible = False
        DataGridView1.Columns("deco_meth_2").Visible = False
        DataGridView1.Columns("RouteDescription").Visible = False

        DataGridView1.Columns("CAT").Width = 100
        DataGridView1.Columns("RouteCatID").Width = 300
        DataGridView1.Columns("RouteID").Width = 500

    End Sub

    Private Sub show_row_in_dgv()
        Try
            Dim i As Integer = 0
            For Each dm As cUS_Exact_Traveler_Route_Accepted In olist_US_Exact_Traveler_Route_Accepted
                DataGridView1.Rows.Add()
                DataGridView1.Item("US_Route_Accepted_ID", i).Value = dm.US_Route_Accepted_ID

                DataGridView1.Item("CAT", i) = createCATCombo()
                DataGridView1.Item("CAT", i).Value = dm.CAT

                'DataGridView1.Item("RouteCatID", i) = createRouteCategoryCombo()
                'If if_RouteCateid_valid(dm.RouteCatID) Then
                DataGridView1.Item("RouteCatID", i).Value = dm.RouteCatID
                'Else
                '    DataGridView1.Item("RouteCatID", i).Value = 0
                'End If


                DataGridView1.Item("RouteCategory", i).Value = dm.RouteCategory

                DataGridView1.Item("deco_meth_1", i).Value = dm.deco_meth_1
                DataGridView1.Item("deco_meth_2", i).Value = dm.deco_meth_2

                'If if_routeCategory_valid(dm.RouteCategory) Then
                '    DataGridView1.Item("RouteID", i) = createRouteCombo(dm.RouteCategory)
                '    DataGridView1.Item("RouteID", i).Value = dm.RouteID
                'End If
                DataGridView1.Item("RouteID", i) = createRouteCombo()
                DataGridView1.Item("RouteID", i).Value = dm.RouteID
                'DataGridView1.Item("RouteDescription", i).Value = dm.RouteDescription
                i = i + 1
            Next
        Catch ex As Exception
            MsgBox("Error in show_row_in_dgv" & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Function createCATCombo() As DataGridViewComboBoxCell
        createCATCombo = New DataGridViewComboBoxCell
        Try

            Dim oCombo As DataGridViewComboBoxCell = New DataGridViewComboBoxCell

            oCombo.DropDownWidth = 100

            oCombo.Items.Add("Ashbury")
            oCombo.Items.Add("Journal")
            oCombo.Items.Add("ORA")
            oCombo.Items.Add("PEN Metal")
            oCombo.Items.Add("PEN Plastic")
            oCombo.Items.Add("PKG")

            createCATCombo = oCombo

        Catch ex As Exception
            MsgBox("Error In createRouteCombo" & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Private Function createRouteCategoryCombo() As DataGridViewComboBoxCell
        createRouteCategoryCombo = New DataGridViewComboBoxCell
        Try

            Dim DB As New cDBA
            Dim strSql As String = ""
            Dim DT As DataTable

            strSql =
                "Select 0 routecatid, '' routecategory union all Select  routecatid, routecategory from EXACT_TRAVELER_Xref_RouteCategory where ISNULL(deco_meth_3, '') = ''  order by routecategory "
            DT = DB.DataTable(strSql)

            Dim oCombo As DataGridViewComboBoxCell = New DataGridViewComboBoxCell

            oCombo.DropDownWidth = 300
            oCombo.DataSource = DT
            oCombo.DisplayMember = "routecategory"
            oCombo.ValueMember = "routecatid"

            createRouteCategoryCombo = oCombo

        Catch ex As Exception
            MsgBox("Error In createRouteCategoryCombo" & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Private Function if_RouteCateid_valid(ByVal in_cateID As Int32) As Boolean
        if_RouteCateid_valid = False
        Try

            Dim DB As New cDBA
            Dim strSql As String = ""
            Dim DT As DataTable

            strSql =
                " Select  routecatid, routecategory from EXACT_TRAVELER_Xref_RouteCategory where ISNULL(deco_meth_3, '') = '' and routecatid = " & in_cateID.ToString
            DT = DB.DataTable(strSql)

            If DT.Rows.Count > 0 Then
                if_RouteCateid_valid = True
            End If

        Catch ex As Exception
            MsgBox("Error in if_RouteCateid_valid" & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Private Function createRouteCombo() As DataGridViewComboBoxCell
        createRouteCombo = New DataGridViewComboBoxCell
        Try

            Dim DB As New cDBA
            Dim strSql As String = ""
            Dim DT As DataTable

            strSql =
                "select 0 routeid , '-' RouteDescription union all select routeid , RouteDescription  from EXACT_TRAVELER_ROUTE where Active = 1  order by routedescription "
            DT = DB.DataTable(strSql)

            Dim oCombo As DataGridViewComboBoxCell = New DataGridViewComboBoxCell

            oCombo.DropDownWidth = 300
            oCombo.DataSource = DT
            oCombo.DisplayMember = "RouteDescription"
            oCombo.ValueMember = "routeid"
            oCombo.MaxDropDownItems = 25



            createRouteCombo = oCombo

        Catch ex As Exception
            MsgBox("Error in createRouteCombo" & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function
    Private Function createRouteCombo(ByVal in_routeCategory As String) As DataGridViewComboBoxCell
        createRouteCombo = New DataGridViewComboBoxCell
        Try

            Dim DB As New cDBA
            Dim strSql As String = ""
            Dim DT As DataTable

            strSql =
                "select  routeid , RouteDescription  from EXACT_TRAVELER_ROUTE where Active = 1 and routecategory = '" & in_routeCategory & "' order by routedescription "
            DT = DB.DataTable(strSql)

            Dim oCombo As DataGridViewComboBoxCell = New DataGridViewComboBoxCell

            oCombo.DropDownWidth = 300
            oCombo.DataSource = DT
            oCombo.DisplayMember = "RouteDescription"
            oCombo.ValueMember = "routeid"

            createRouteCombo = oCombo

        Catch ex As Exception
            MsgBox("Error in createRouteCombo" & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Private Function if_routeCategory_valid(ByVal in_routeCategory As String) As Boolean
        if_routeCategory_valid = False
        Try

            Dim DB As New cDBA
            Dim strSql As String = ""
            Dim DT As DataTable

            strSql =
                " Select 1 from EXACT_TRAVELER_ROUTE where Active = 1 and routecategory = '" & in_routeCategory & "'"
            DT = DB.DataTable(strSql)

            If DT.Rows.Count > 0 Then
                if_routeCategory_valid = True
            End If

        Catch ex As Exception
            MsgBox("Error in if_RouteCateid_valid" & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim _cUS_Exact_Traveler_Route_Accepted As cUS_Exact_Traveler_Route_Accepted = olist_US_Exact_Traveler_Route_Accepted.Find(Function(j As cUS_Exact_Traveler_Route_Accepted) j.US_Route_Accepted_ID = DataGridView1.Item("US_Route_Accepted_ID", e.RowIndex).Value)
        If _cUS_Exact_Traveler_Route_Accepted IsNot Nothing Then
            Select Case DataGridView1.Columns(e.ColumnIndex).Name
                Case "CAT"
                    _cUS_Exact_Traveler_Route_Accepted.CAT = DataGridView1.Item("CAT", e.RowIndex).Value
                Case "RouteCatID"
                    _cUS_Exact_Traveler_Route_Accepted.RouteCatID = DataGridView1.Item("RouteCatID", e.RowIndex).Value
                    _cUS_Exact_Traveler_Route_Accepted.RouteCategory = DataGridView1.Item("RouteCatID", e.RowIndex).FormattedValue
                Case "RouteID"
                    _cUS_Exact_Traveler_Route_Accepted.RouteID = DataGridView1.Item("RouteID", e.RowIndex).Value
                    _cUS_Exact_Traveler_Route_Accepted.RouteDescription = DataGridView1.Item("RouteID", e.RowIndex).FormattedValue
            End Select
            _cUS_Exact_Traveler_Route_Accepted_DAL.Save(_cUS_Exact_Traveler_Route_Accepted)
        Else
            If DataGridView1.Item("US_Route_Accepted_ID", e.RowIndex).Value > 0 Then
                Dim cUS_Exact_Traveler_Route_Accepted As cUS_Exact_Traveler_Route_Accepted = New cUS_Exact_Traveler_Route_Accepted()
                Select Case DataGridView1.Columns(e.ColumnIndex).Name
                    Case "CAT"
                        If DataGridView1.Item("CAT", e.RowIndex).Value = "" Then
                            MsgBox("Please select Cat column!")
                            Return
                        End If
                        cUS_Exact_Traveler_Route_Accepted.CAT = DataGridView1.Item("CAT", e.RowIndex).Value
                    Case "RouteID"
                        cUS_Exact_Traveler_Route_Accepted.RouteID = DataGridView1.Item("RouteID", e.RowIndex).Value
                        cUS_Exact_Traveler_Route_Accepted.RouteDescription = DataGridView1.Item("RouteID", e.RowIndex).FormattedValue

                        cUS_Exact_Traveler_Route_Accepted.CAT = DataGridView1.Item("CAT", e.RowIndex).Value
                End Select
                _cUS_Exact_Traveler_Route_Accepted_DAL.Save(cUS_Exact_Traveler_Route_Accepted)

            End If

        End If

    End Sub

    Private Sub Button_ADD_Click(sender As Object, e As EventArgs) Handles Button_ADD.Click
        Try
            DataGridView1.Rows.Add()
            Dim i As Int32 = DataGridView1.Rows.Count - 1
            DataGridView1.Item("US_Route_Accepted_ID", i).Value = _cUS_Exact_Traveler_Route_Accepted_DAL.HighId()

            DataGridView1.Item("CAT", i) = createCATCombo()
            DataGridView1.Item("CAT", i).Value = "Ashbury"

            DataGridView1.Item("RouteID", i) = createRouteCombo()
            DataGridView1.Item("RouteID", i).Value = 0

            DataGridView1.Item("RouteCatID", i).Value = ""
            DataGridView1.Item("RouteCategory", i).Value = ""
            DataGridView1.Item("deco_meth_1", i).Value = ""
            DataGridView1.Item("deco_meth_2", i).Value = ""
            DataGridView1.Item("RouteDescription", i).Value = ""

            DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(1)
        Catch ex As Exception
            MsgBox("Error in Button_ADD_Click" & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Debug.Print("column" & e.ColumnIndex.ToString)
        Debug.Print("row" & e.RowIndex.ToString)
    End Sub

    Private Sub ButtonDEL_Click(sender As Object, e As EventArgs) Handles ButtonDEL.Click

        Dim _US_Route_Accepted_ID As Int32 = DataGridView1.Item("US_Route_Accepted_ID", DataGridView1.CurrentRow.Index).Value
        Debug.Print("rowindex" & DataGridView1.CurrentRow.Index.ToString & "-----" & _US_Route_Accepted_ID.ToString)
        _cUS_Exact_Traveler_Route_Accepted_DAL.Delete(_US_Route_Accepted_ID)

        dgvshow()
    End Sub

    Private Sub ButtonRefresh_Click(sender As Object, e As EventArgs) Handles ButtonRefresh.Click
        dgvshow()
    End Sub

    'Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click
    '    Dim i As Int32

    '    For i = 0 To DataGridView1.RowCount() - 1
    '        'Check condition
    '        If DataGridView1.Rows(i).Cells("CAT").
    '        End If
    '    Next
    'End Sub
End Class