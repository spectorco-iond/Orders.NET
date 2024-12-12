Public Class cOEFlashCustomer

    ' Unexposed properties
    Private m_ClassName As String = "cOEFlashCustomer"
    Private db As New cDBA

    ' Exposed properties
    Private m_intCustomer_ID As Integer
    Private m_strCus_No As String
    Private m_intEQP_Type As Integer
    Private m_intEQP_Plus_Pct As Integer
    Private m_datEQP_Start_Date As Date
    Private m_datEQP_End_Date As Date

#Region "Public constructors ##############################################"

    Public Sub New()

        Call Init()

    End Sub

    Public Sub New(ByVal pstrCus_No As String)

        Call Init()
        m_strCus_No = pstrCus_No

    End Sub

    Private Sub New(ByVal pintCustomer_ID As Integer)

        m_intCustomer_ID = pintCustomer_ID
        Call Load()

    End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_intCustomer_ID = 0

            m_strCus_No = ""
            m_intEQP_Type = 0
            m_intEQP_Plus_Pct = 0
            m_datEQP_Start_Date = NoDate()
            m_datEQP_End_Date = NoDate()

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Load()

        Try

            Dim dt As DataTable

            Dim strSql As String = _
            "SELECT * " & _
            "FROM   MDB_CFG_CUSTOMER " & _
            "WHERE  CUSTOMER_ID = " & m_intCustomer_ID.ToString

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                Call LoadLine(dt.Rows(0))
            Else
                Call Init()
            End If

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            With pdrRow

                If Not (.Item("Customer_ID").Equals(DBNull.Value)) Then m_intCustomer_ID = .Item("Customer_ID")
                If Not (.Item("Cus_No").Equals(DBNull.Value)) Then m_strCus_No = .Item("Cus_No").ToString
                If Not (.Item("EQP_Type").Equals(DBNull.Value)) Then m_intEQP_Type = .Item("EQP_Type")
                If Not (.Item("EQP_Plus_Pct").Equals(DBNull.Value)) Then m_intEQP_Plus_Pct = .Item("EQP_Plus_Pct")
                If Not (.Item("EQP_Start_Date").Equals(DBNull.Value)) Then m_datEQP_Start_Date = .Item("EQP_Start_Date")
                If Not (.Item("EQP_End_Date").Equals(DBNull.Value)) Then m_datEQP_End_Date = .Item("EQP_End_Date")

            End With

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Reload()

        Call Reset()
        Call Load()

    End Sub

    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            pdrRow.Item("Cus_No") = m_strCus_No

            ' Never save m_intCharge_Usage_ID as it is the incremental ID
            ' pdrRow.Item("Customer_ID") = m_intCustomer_ID

            pdrRow.Item("EQP_Type") = m_intEQP_Type
            pdrRow.Item("EQP_Plus_Pct") = m_intEQP_Plus_Pct
            If m_datEQP_Start_Date.Year <> 1 Then pdrRow.Item("EQP_Start_Date") = m_datEQP_Start_Date
            If m_datEQP_End_Date.Year <> 1 Then pdrRow.Item("EQP_End_Date") = m_datEQP_End_Date

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    Public Sub Delete()

        Try
            If m_intCustomer_ID = 0 Then Exit Sub

            Dim dt As New DataTable

            Dim strSql As String = _
            "DELETE FROM MDB_CFG_CUSTOMER " & _
            "WHERE  CUSTOMER_ID = " & m_intCustomer_ID & " "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Reset()

        Try

            m_strCus_No = String.Empty
            m_intEQP_Type = 0
            m_intEQP_Plus_Pct = 0
            m_datEQP_Start_Date = NoDate()
            m_datEQP_End_Date = NoDate()

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Save()

        Try

            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String = _
            "SELECT * " & _
            "FROM   MDB_CFG_CUSTOMER " & _
            "WHERE  CUSTOMER_ID = " & m_intCustomer_ID & " "

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

            If m_intCustomer_ID <> 0 Then m_intCustomer_ID = db.DataTable("SELECT MAX(Customer_ID) AS Max_Customer_ID FROM MDB_CFG_CUSTOMER WITH (Nolock)").Rows(0).Item("Max_Customer_ID")

        Catch er As Exception
            MsgBox("Error in " & m_ClassName & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property Customer_ID() As Integer
        Get
            Customer_ID = m_intCustomer_ID
        End Get
        Set(ByVal value As Integer)
            m_intCustomer_ID = value
        End Set
    End Property
    Public Property Cus_No() As String
        Get
            Cus_No = m_strCus_No
        End Get
        Set(ByVal value As String)
            m_strCus_No = value
        End Set
    End Property
    Public Property EQP_Type() As Integer
        Get
            EQP_Type = m_intEQP_Type
        End Get
        Set(ByVal value As Integer)
            m_intEQP_Type = value
        End Set
    End Property
    Public Property EQP_Plus_Pct() As Integer
        Get
            EQP_Plus_Pct = m_intEQP_Plus_Pct
        End Get
        Set(ByVal value As Integer)
            m_intEQP_Plus_Pct = value
        End Set
    End Property
    Public Property EQP_Start_Date() As Date
        Get
            EQP_Start_Date = m_datEQP_Start_Date
        End Get
        Set(ByVal value As Date)
            m_datEQP_Start_Date = value
        End Set
    End Property
    Public Property EQP_End_Date() As Date
        Get
            EQP_End_Date = m_datEQP_End_Date
        End Get
        Set(ByVal value As Date)
            m_datEQP_End_Date = value
        End Set
    End Property

#End Region

End Class
