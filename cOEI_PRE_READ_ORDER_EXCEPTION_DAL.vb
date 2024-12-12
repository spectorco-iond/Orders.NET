Imports System.Data.Odbc
Imports System.Collections.Generic

Public Class cOEI_PRE_READ_ORDER_EXCEPTION_DAL
    Public Function Load_By_OrderNo(ByVal in_Order_no As String) As List(Of cOEI_PRE_READ_ORDER_EXCEPTION)
        Dim myList As List(Of cOEI_PRE_READ_ORDER_EXCEPTION) = New List(Of cOEI_PRE_READ_ORDER_EXCEPTION)
        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "Select Case oe.*,mc.ENUM_VALUE  " &
            "From OEI_PRE_READ_ORDER_EXCEPTION oe  " &
            "inner Join mdb_cfg_enum mc on oe.Exception_type = mc.ID And mc.ENUM_CAT = 'OEI_PRE_READ_ORDER_EXCEPTION' " &
            "WHERE  ORDER_NO = '" & in_Order_no & "' "

            dt = db.DataTable(strSql)

            Dim i As Int32
            If dt.Rows.Count <> 0 Then
                For i = 0 To dt.Rows.Count - 1
                    Dim oEnum = New cOEI_PRE_READ_ORDER_EXCEPTION
                    Call LoadLine(oEnum, dt.Rows(i))
                    myList.Add(oEnum)
                Next
            End If

        Catch ex As Exception
            MsgBox("Error in cOEI_PRE_READ_ORDER_EXCEPTION_DAL." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
        Return myList
    End Function

    Public Function Load_EXCEPTION_Enum_list() As List(Of cMdb_Cfg_Enum)
        Dim myList As List(Of cMdb_Cfg_Enum) = New List(Of cMdb_Cfg_Enum)

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "select * from mdb_cfg_enum where ENUM_CAT = 'OEI_PRE_READ_ORDER_EXCEPTION' "

            dt = db.DataTable(strSql)
            Dim oEnum_all = New cMdb_Cfg_Enum

            oEnum_all.Id = "0"
            oEnum_all.Enum_Cat = "-ALL-"
            oEnum_all.Enum_Value = "-ALL-"
            oEnum_all.Enum_Sub_Cat = "-ALL-"

            myList.Add(oEnum_all)

            Dim i As Int32
            If dt.Rows.Count <> 0 Then
                For i = 0 To dt.Rows.Count - 1
                    Dim oEnum = New cMdb_Cfg_Enum
                    If Not (dt.Rows(i).Item("Id").Equals(DBNull.Value)) Then oEnum.Id = dt.Rows(i).Item("Id")
                    If Not (dt.Rows(i).Item("Enum_Cat").Equals(DBNull.Value)) Then oEnum.Enum_Cat = dt.Rows(i).Item("Enum_Cat").ToString
                    If Not (dt.Rows(i).Item("Enum_Value").Equals(DBNull.Value)) Then oEnum.Enum_Value = dt.Rows(i).Item("Enum_Value").ToString
                    If Not (dt.Rows(i).Item("Enum_Sub_Cat").Equals(DBNull.Value)) Then oEnum.Enum_Sub_Cat = dt.Rows(i).Item("Enum_Sub_Cat").ToString
                    If Not (dt.Rows(i).Item("User_Login").Equals(DBNull.Value)) Then oEnum.User_Login = dt.Rows(i).Item("User_Login").ToString
                    If Not (dt.Rows(i).Item("Create_Ts").Equals(DBNull.Value)) Then oEnum.Create_Ts = dt.Rows(i).Item("Create_Ts")
                    If Not (dt.Rows(i).Item("Update_Ts").Equals(DBNull.Value)) Then oEnum.Update_Ts = dt.Rows(i).Item("Update_Ts")

                    '
                    myList.Add(oEnum)
                Next
            End If

        Catch ex As Exception
            MsgBox("Error in cOEI_PRE_READ_ORDER_EXCEPTION_DAL." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
        Return myList
    End Function

    Public Function Load_EXCEPTION_report_ALL() As DataTable


        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "select n.ENUM_VALUE,e.order_no,e.oe_po_no,e.item_cd,e.item_no,e.ENTER_person, " _
& " e.CREATE_TS from OEI_PRE_READ_ORDER_EXCEPTION e " _
& "  inner join mdb_cfg_enum n On e.Exception_type = n.id And ENUM_CAT = 'OEI_PRE_READ_ORDER_EXCEPTION' "

            dt = db.DataTable(strSql)
            Return dt
        Catch ex As Exception
            MsgBox("Error in cOEI_PRE_READ_ORDER_EXCEPTION_DAL." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function
    Public Function Load_EXCEPTION_report_by_Id(ByVal enum_id As Int16) As DataTable


        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql =
            "select n.ENUM_VALUE,e.order_no,e.oe_po_no,e.item_cd,e.item_no,e.ENTER_person, " _
& " e.CREATE_TS from OEI_PRE_READ_ORDER_EXCEPTION e " _
& "  inner join mdb_cfg_enum n On e.Exception_type = n.id And ENUM_CAT = 'OEI_PRE_READ_ORDER_EXCEPTION' " _
& " where n.ID = " & enum_id.ToString() & " "

            dt = db.DataTable(strSql)
            Return dt
        Catch ex As Exception
            MsgBox("Error in cOEI_PRE_READ_ORDER_EXCEPTION_DAL." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function
    Private Sub LoadLine(ByVal pClass As cOEI_PRE_READ_ORDER_EXCEPTION, ByRef pdrRow As DataRow)

        Try

            If Not (pdrRow.Item("ID").Equals(DBNull.Value)) Then pClass.Id = pdrRow.Item("ID")
            If Not (pdrRow.Item("ORDER_NO").Equals(DBNull.Value)) Then pClass.ORDER_NO = pdrRow.Item("ORDER_NO").ToString
            If Not (pdrRow.Item("OE_PO_NO").Equals(DBNull.Value)) Then pClass.OE_PO_NO = pdrRow.Item("OE_PO_NO").ToString
            If Not (pdrRow.Item("ITEM_CD").Equals(DBNull.Value)) Then pClass.ITEM_CD = pdrRow.Item("ITEM_CD").ToString
            If Not (pdrRow.Item("ITEM_NO").Equals(DBNull.Value)) Then pClass.ITEM_NO = pdrRow.Item("ITEM_NO").ToString
            If Not (pdrRow.Item("Exception_type").Equals(DBNull.Value)) Then pClass.Exception_type = pdrRow.Item("Exception_type").ToString
            If Not (pdrRow.Item("timestamp").Equals(DBNull.Value)) Then pClass.timestamp = pdrRow.Item("timestamp").ToString
            If Not (pdrRow.Item("ENTER_person").Equals(DBNull.Value)) Then pClass.ENTER_person = pdrRow.Item("ENTER_person").ToString
            If Not (pdrRow.Item("USERID").Equals(DBNull.Value)) Then pClass.USERID = pdrRow.Item("USERID").ToString
            If Not (pdrRow.Item("USER_LOGIN").Equals(DBNull.Value)) Then pClass.USER_LOGIN = pdrRow.Item("USER_LOGIN").ToString
            If Not (pdrRow.Item("Create_Ts").Equals(DBNull.Value)) Then pClass.Create_Ts = pdrRow.Item("Create_Ts")
            If Not (pdrRow.Item("Update_Ts").Equals(DBNull.Value)) Then pClass.Update_Ts = pdrRow.Item("Update_Ts")
            If Not (pdrRow.Item("ENUM_VALUE").Equals(DBNull.Value)) Then pClass.Exception_text = pdrRow.Item("ENUM_VALUE").ToString

        Catch ex As Exception
            MsgBox("Error in cOEI_PRE_READ_ORDER_EXCEPTION_DAL." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub

    Public Function Save(ByVal pClass As cOEI_PRE_READ_ORDER_EXCEPTION) As Boolean
        Save = False
        Try
            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow
            Dim strSql As String

            strSql =
                "SELECT * " &
                "FROM   OEI_PRE_READ_ORDER_EXCEPTION " &
            "WHERE  Id = " & pClass.Id & " "
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
            MsgBox("Error in cMdb_Cfg_Enum." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

    Private Sub SaveLine(ByRef pClass As cOEI_PRE_READ_ORDER_EXCEPTION, ByRef pdrRow As DataRow)
        Try
            'pdrRow.Item("Id") = pClass.Id
            pdrRow.Item("ORDER_NO") = pClass.ORDER_NO
            pdrRow.Item("OE_PO_NO") = pClass.OE_PO_NO

            pdrRow.Item("ITEM_CD") = pClass.ITEM_CD
            pdrRow.Item("ITEM_NO") = pClass.ITEM_NO
            pdrRow.Item("Exception_type") = pClass.Exception_type
            pdrRow.Item("timestamp") = pClass.timestamp
            pdrRow.Item("ENTER_person") = pClass.ENTER_person
            pdrRow.Item("USERID") = pClass.USERID
            pdrRow.Item("USER_LOGIN") = pClass.USER_LOGIN
            pdrRow.Item("Create_Ts") = pClass.Create_Ts
            pdrRow.Item("Update_Ts") = pClass.Update_Ts
        Catch ex As Exception
            MsgBox("Error in cMdb_Cfg_Enum." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Sub
End Class
