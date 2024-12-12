Public Class cMdb_Cus_Dec_Supp_Item

    Private m_Cus_Dec_Supp_Item_ID As Integer = 0
    Private m_Cus_Dec_ID As Integer
    Private m_Item_No As String
    Private m_Qty_Ordered As Double
    Private m_Unit_Price As Double
    Private m_User_Login As String
    Private m_Update_TS As DateTime
    Private m_DataTable As DataTable

#Region "Public constructors ##############################################"

    'Public Sub New()
    Public Sub New()

        Call Init()

    End Sub

    'Public Sub New(ByVal pID As Integer)
    Public Sub New(ByVal pCus_Dec_Supp_Item_ID As Integer)

        Try

            Call Init()

            m_Cus_Dec_Supp_Item_ID = pCus_Dec_Supp_Item_ID

            Call Load(pCus_Dec_Supp_Item_ID)

        Catch er As Exception
            MsgBox("Error in cMDB_CUS_DEC_SUPP_ITEM." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_Cus_Dec_Supp_Item_ID = 0

            Call Reset()

        Catch er As Exception
            MsgBox("Error in cMDB_CUS_DEC_SUPP_ITEM." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Load is made private, it is loaded automatically on New.
    Private Sub Load(ByVal pintCus_Dec_Supp_Item_ID As Integer)

        Try
            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = _
            "SELECT * " & _
            "FROM   MDB_CUS_DEC_SUPP_ITEM PT WITH (Nolock) " & _
            "WHERE  PT.Cus_Dec_Supp_Item_ID = " & pintCus_Dec_Supp_Item_ID & " "

            dt = db.DataTable(strSql)

            Call LoadLine(dt.Rows(0))

        Catch er As Exception
            MsgBox("Error in cMDB_CUS_DEC_SUPP_ITEM." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Load is made private, it is loaded automatically on New.
    ' Loads the Comment fields from the DataRow record
    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            If Not (pdrRow.Item("Cus_Dec_Supp_Item_ID").Equals(DBNull.Value)) Then m_Cus_Dec_Supp_Item_ID = pdrRow.Item("Cus_Dec_Supp_Item_ID")
            If Not (pdrRow.Item("Cus_Dec_ID").Equals(DBNull.Value)) Then m_Cus_Dec_ID = pdrRow.Item("Cus_Dec_ID")
            If Not (pdrRow.Item("Item_No").Equals(DBNull.Value)) Then m_Item_No = pdrRow.Item("Item_No").ToString
            If Not (pdrRow.Item("Qty_Ordered").Equals(DBNull.Value)) Then m_Qty_Ordered = pdrRow.Item("Qty_Ordered")
            If Not (pdrRow.Item("Unit_Price").Equals(DBNull.Value)) Then m_Unit_Price = pdrRow.Item("Unit_Price")
            If Not (pdrRow.Item("User_Login").Equals(DBNull.Value)) Then m_User_Login = pdrRow.Item("User_Login").ToString
            If Not (pdrRow.Item("Update_TS").Equals(DBNull.Value)) Then m_Update_TS = pdrRow.Item("Update_TS")

        Catch er As Exception
            MsgBox("Error in cMDB_CUS_DEC_SUPP_ITEM." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Saves the Comment fields into the DataRow record
    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            pdrRow.Item("Cus_Dec_ID") = m_Cus_Dec_ID
            pdrRow.Item("Item_No") = m_Item_No
            pdrRow.Item("Qty_Ordered") = m_Qty_Ordered
            pdrRow.Item("Unit_Price") = m_Unit_Price
            pdrRow.Item("User_Login") = m_User_Login
            pdrRow.Item("Update_TS") = Date.Now()

        Catch er As Exception
            MsgBox("Error in cMDB_CUS_DEC_SUPP_ITEM." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    ' Deletes the current Comment from the database, if it exists.
    Public Sub Delete()

        Try

            If m_Cus_Dec_Supp_Item_ID = 0 Then Exit Sub

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            strSql = _
            "DELETE FROM MDB_CUS_DEC_SUPP_ITEM " & _
            "WHERE  Cus_Dec_Supp_Item_ID = " & m_Cus_Dec_Supp_Item_ID & " "

            db.Execute(strSql)

        Catch er As Exception
            MsgBox("Error in cMDB_CUS_DEC_SUPP_ITEM." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Reset every non mandatory field of a Comment
    Public Sub Reset()

        Try
            ' Resets the locals to empty values, all but ID, Cus_Dec_ID and Qty_Ordered
            m_Cus_Dec_ID = 0
            m_Item_No = String.Empty
            m_Qty_Ordered = 0
            m_Unit_Price = 0
            m_User_Login = String.Empty
            m_Update_TS = Nothing

        Catch er As Exception
            MsgBox("Error in cMDB_CUS_DEC_SUPP_ITEM." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Update the current Comment into the database, or creates it if not existing
    Public Sub Save()

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String

            strSql = _
            "SELECT * " & _
            "FROM   MDB_CUS_DEC_SUPP_ITEM " & _
            "WHERE  Cus_Dec_Supp_Item_ID = " & m_Cus_Dec_Supp_Item_ID & " "

            dt = db.DataTable(strSql)

            If dt.Rows.Count = 0 Then
                drRow = dt.NewRow()
            Else
                drRow = dt.Rows(0)
            End If
            'drRow = IIf(dt.Rows.Count = 0, dt.NewRow(), dt.Rows(0))

            Call SaveLine(drRow)

            If dt.Rows.Count = 0 Then
                db.DBDataTable.Rows.Add(drRow)
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
            Else
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
                db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            End If

            db.DBDataAdapter.Update(db.DBDataTable)

        Catch er As Exception
            MsgBox("Error in cMDB_CUS_DEC_SUPP_ITEM." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property Cus_Dec_Supp_Item_ID() As Integer
        Get
            Cus_Dec_Supp_Item_ID = m_Cus_Dec_Supp_Item_ID
        End Get
        Set(ByVal value As Integer)
            m_Cus_Dec_Supp_Item_ID = value
        End Set
    End Property
    Public Property Cus_Dec_ID() As Integer
        Get
            Cus_Dec_ID = m_Cus_Dec_ID
        End Get
        Set(ByVal value As Integer)
            m_Cus_Dec_ID = value
        End Set
    End Property
    Public Property Item_No() As String
        Get
            Item_No = m_Item_No
        End Get
        Set(ByVal value As String)
            m_Item_No = value
        End Set
    End Property
    Public Property Unit_Price() As Double
        Get
            Unit_Price = m_Unit_Price
        End Get
        Set(ByVal value As Double)
            m_Unit_Price = value
        End Set
    End Property
    Public Property Qty_Ordered() As Double
        Get
            Qty_Ordered = m_Qty_Ordered
        End Get
        Set(ByVal value As Double)
            m_Qty_Ordered = value
        End Set
    End Property
    Public Property User_Login() As String
        Get
            User_Login = m_User_Login
        End Get
        Set(ByVal value As String)
            m_User_Login = value
        End Set
    End Property
    Public Property Update_TS() As DateTime
        Get
            Update_TS = m_Update_TS
        End Get
        Set(ByVal value As DateTime)
            m_Update_TS = value
        End Set
    End Property

#End Region

End Class
