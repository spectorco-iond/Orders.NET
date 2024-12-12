Imports System.Data.Odbc 

Public Class cOei_Address_Book_DAL


  Public Function Load(ByVal pintID As Integer) As cOei_Address_Book

      Load = New cOei_Address_Book() 

      Try

          Dim db As New cDBA
          Dim dt As New DataTable

          Dim strSql As String
          strSql = _
            "SELECT * " & _
            "FROM   Oei_Address_Book WITH (Nolock) " & _
            "WHERE  Id = " & pintID & " "

          dt = db.DataTable(strSql)

          If dt.Rows.Count <> 0 Then
              Call LoadLine(Load, dt.Rows(0))
          End If 

        Catch ex as Exception
            MsgBox("Error in cOei_Address_Book." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

 

    Public Function LoadAll() As DataTable
        Dim db As New cDBA
        Dim dt As New DataTable
        Try
            Dim strSql As String

            strSql = " Select * from Oei_Address_Book  "

            dt = db.DataTable(strSql)

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
        Return dt
    End Function

    Private Sub LoadLine(pClass As cOei_Address_Book, ByRef pdrRow As DataRow)

        Try


        Catch ex As Exception
            MsgBox("Error in cOei_Address_Book." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    ' Insert/Update the current row into the database, or creates it if not existing
    Public Function Save(pClass As cOei_Address_Book) As Boolean

        Save = False

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String

            strSql = _
                "SELECT * " & _
                "FROM   Oei_Address_Book " & _
            "WHERE  Id = " & pClass.addrBook_Id & " "

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
            MsgBox("Error in cOei_Address_Book." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function


    Private Sub SaveLine(ByRef pClass As cOei_Address_Book, ByRef pdrRow As DataRow)

        Try

            pdrRow.Item("ID") = pClass.addrBook_Id
            pdrRow.Item("FullName") = pClass.addrBook_Fullname
            pdrRow.Item("Mail") = pClass.addrBook_Mail
        Catch ex As Exception
            MsgBox("Error in cMdb_Cfg_Imp_Loc." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Sub


    ' Deletes the current row from the database, if it exists.
    Public Function Delete(pintID As Integer) As Boolean

        Delete = False

        Try

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            strSql = _
                "DELETE FROM Oei_Address_Book " & _
            "WHERE  Id = " & pintID & " "

            db.Execute(strSql)

            Delete = True

        Catch ex As Exception
            MsgBox("Error in cOei_Address_Book." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Function

End Class ' cOei_Address_Book


