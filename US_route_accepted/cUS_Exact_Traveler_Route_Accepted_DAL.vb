Imports System.Data.Odbc
Imports System.Collections.Generic

Public Class cUS_Exact_Traveler_Route_Accepted_DAL

    Public Function Load(ByVal pintID As Integer) As cUS_Exact_Traveler_Route_Accepted

        Load = New cUS_Exact_Traveler_Route_Accepted()

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
              "SELECT US_Route_Accepted_ID,CAT,isnull(RouteCatID, 0) RouteCatID,isnull( RouteCategory,'') RouteCategory,deco_meth_1,deco_meth_2,RouteID,RouteDescription " &
              "FROM   US_Exact_Traveler_Route_Accepted_table WITH (Nolock) " &
              "WHERE  US_Route_Accepted_ID = " & pintID &
              " and ( exists(select 1 from EXACT_TRAVELER_Xref_RouteCategory x where ISNULL(x.deco_meth_3, '') = '' and x.RouteCatID = US_Exact_Traveler_Route_Accepted_table.RouteCatID )  or isnull(RouteCatID, 0) = 0 )" &
              " and exists(select 1 from EXACT_TRAVELER_ROUTE r where r.Routeid = US_Exact_Traveler_Route_Accepted_table.RouteID ) order by CAT,RouteDescription"


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(Load, dt.Rows(0))
            End If

        Catch ex As Exception
            MsgBox("Error in cUS_Exact_Traveler_Route_Accepted_DAL." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

    Private Sub LoadLine(pClass As cUS_Exact_Traveler_Route_Accepted, ByRef pdrRow As DataRow)
        Try
            If Not (pdrRow.Item("US_Route_Accepted_ID").Equals(DBNull.Value)) Then pClass.US_Route_Accepted_ID = pdrRow.Item("US_Route_Accepted_ID")
            If Not (pdrRow.Item("CAT").Equals(DBNull.Value)) Then pClass.CAT = pdrRow.Item("CAT")
            If Not (pdrRow.Item("RouteCatID").Equals(DBNull.Value)) Then pClass.RouteCatID = pdrRow.Item("RouteCatID")
            If Not (pdrRow.Item("RouteCategory").Equals(DBNull.Value)) Then pClass.RouteCategory = pdrRow.Item("RouteCategory").ToString
            If Not (pdrRow.Item("deco_meth_1").Equals(DBNull.Value)) Then pClass.deco_meth_1 = pdrRow.Item("deco_meth_1").ToString
            If Not (pdrRow.Item("deco_meth_2").Equals(DBNull.Value)) Then pClass.deco_meth_2 = pdrRow.Item("deco_meth_2").ToString
            If Not (pdrRow.Item("RouteID").Equals(DBNull.Value)) Then pClass.RouteID = pdrRow.Item("RouteID")
            If Not (pdrRow.Item("RouteDescription").Equals(DBNull.Value)) Then pClass.RouteDescription = pdrRow.Item("RouteDescription")

        Catch ex As Exception
            MsgBox("Error in cUS_Exact_Traveler_Route_Accepted_DAL." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub

    Public Function GetAllList() As List(Of cUS_Exact_Traveler_Route_Accepted)

        GetAllList = New List(Of cUS_Exact_Traveler_Route_Accepted)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim myList = New List(Of cUS_Exact_Traveler_Route_Accepted)
            Dim i As Integer
            Dim strSql As String

            strSql =
              "SELECT US_Route_Accepted_ID,CAT,isnull(RouteCatID,0) RouteCatID, isnull( RouteCategory,'') RouteCategory,deco_meth_1,deco_meth_2,RouteID,RouteDescription " &
            " FROM   US_Exact_Traveler_Route_Accepted_table WITH (Nolock) " &
              " where " &
              " ( exists(select 1 from EXACT_TRAVELER_Xref_RouteCategory x where ISNULL(x.deco_meth_3, '') = '' and x.RouteCatID = US_Exact_Traveler_Route_Accepted_table.RouteCatID )  or isnull(RouteCatID, 0) = 0 )" &
              " and exists(select 1 from EXACT_TRAVELER_ROUTE r where r.Routeid = US_Exact_Traveler_Route_Accepted_table.RouteID ) order by CAT,RouteDescription"


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                For i = 0 To dt.Rows.Count - 1
                    Dim oEnum = New cUS_Exact_Traveler_Route_Accepted
                    Call LoadLine(oEnum, dt.Rows(i))
                    myList.Add(oEnum)
                Next
            End If

            GetAllList = myList

        Catch ex As Exception
            MsgBox("Error in cMdb_Item_Imp_Loc." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

    Public Function LoadAll() As DataTable
        Dim db As New cDBA
        Dim dt As New DataTable
        Try
            Dim strSql As String

            strSql = "Select US_Route_Accepted_ID,CAT,isnull(RouteCatID,0) RouteCatID,isnull( RouteCategory,'') RouteCategory,deco_meth_1,deco_meth_2,RouteID,RouteDescription " &
            " from US_Exact_Traveler_Route_Accepted_table " &
              " where " &
              " ( exists(select 1 from EXACT_TRAVELER_Xref_RouteCategory x where ISNULL(x.deco_meth_3, '') = '' and x.RouteCatID = US_Exact_Traveler_Route_Accepted_table.RouteCatID )  or isnull(RouteCatID, 0) = 0 )" &
              " and exists(select 1 from EXACT_TRAVELER_ROUTE r where r.Routeid = US_Exact_Traveler_Route_Accepted_table.RouteID ) order by CAT,RouteDescription"

            dt = db.DataTable(strSql)

        Catch er As Exception
            MsgBox("Error in cUS_Exact_Traveler_Route_Accepted_DAL." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
        Return dt
    End Function

    ' Insert/Update the current row into the database, or creates it if not existing
    Public Function Save(ByVal pClass As cUS_Exact_Traveler_Route_Accepted) As Boolean
        Save = False
        Try
            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow
            Dim strSql As String

            strSql =
                "SELECT US_Route_Accepted_ID,CAT,RouteCatID,RouteCategory,deco_meth_1,deco_meth_2,RouteID,RouteDescription " &
                "FROM   US_Exact_Traveler_Route_Accepted_table " &
            "WHERE  US_Route_Accepted_ID = " & pClass.US_Route_Accepted_ID & " "

            dt = db.DataTable(strSql)

            If dt.Rows.Count = 0 Then
                drRow = dt.NewRow()
            Else
                drRow = dt.Rows(0)
            End If

            Call SaveLine(pClass, drRow)

            If dt.Rows.Count = 0 Then
                db.DBDataTable.Rows.Add(drRow)
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
            Else
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            End If

            db.DBDataAdapter.Update(db.DBDataTable)

        Catch ex As Exception
            MsgBox("Error in cUS_Exact_Traveler_Route_Accepted_DAL." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

    Private Sub SaveLine(ByRef pClass As cUS_Exact_Traveler_Route_Accepted, ByRef pdrRow As DataRow)

        Try
            pdrRow.Item("US_Route_Accepted_ID") = pClass.US_Route_Accepted_ID
            pdrRow.Item("CAT") = pClass.CAT
            pdrRow.Item("RouteCatID") = pClass.RouteCatID
            pdrRow.Item("RouteCategory") = pClass.RouteCategory
            pdrRow.Item("deco_meth_1") = pClass.deco_meth_1
            pdrRow.Item("deco_meth_2") = pClass.deco_meth_2
            pdrRow.Item("RouteID") = pClass.RouteID
            pdrRow.Item("RouteDescription") = pClass.RouteDescription

        Catch ex As Exception
            MsgBox("Error in cUS_Exact_Traveler_Route_Accepted." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub

    ' Deletes the current row from the database, if it exists.
    Public Function Delete(ByVal pintID As Integer) As Boolean

        Delete = False

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            strSql =
                "DELETE FROM US_Exact_Traveler_Route_Accepted_table " &
            "WHERE  US_Route_Accepted_ID = " & pintID & " "

            db.Execute(strSql)

            Delete = True

        Catch ex As Exception
            MsgBox("Error in cUS_Exact_Traveler_Route_Accepted." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

    Public Function HighId() As Integer
        Dim db As New cDBA
        Dim dt As New DataTable
        Try
            Dim strSql As String

            strSql = "select max(US_Route_Accepted_ID) as 'Value' from US_Exact_Traveler_Route_Accepted_table"

            dt = db.DataTable(strSql)

        Catch er As Exception
            MsgBox("Error in HighId " & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
        If IsDBNull(dt.Rows(0).Item("Value")) Then
            Return 0
        Else
            Return dt.Rows(0).Item("Value") + 1
        End If
    End Function

End Class ' cUS_Exact_Traveler_Route_Accepted


