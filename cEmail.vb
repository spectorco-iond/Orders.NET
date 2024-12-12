Public Class cEmail

    Private m_ID As Integer = 0
    Private m_Email_From As String
    Private m_Email_To As String
    Private m_Email_To_Name As String
    Private m_Email_CC As String
    Private m_Email_CC_Name As String
    Private m_Subject As String
    Private m_Body As String
    Private m_Ord_Guid As String
    Private m_Ord_No As String
    Private m_CreateTS As Date
    Private m_SendTS As Date
    Private m_UserID As String

#Region "Public constructors ##############################################"

    'Public Sub New()
    Public Sub New()

        Call Init()

    End Sub

    'Public Sub New(ByVal pID As Integer)
    Public Sub New(ByVal pID As Integer)

        Try

            Call Init()

            m_ID = pID

            Call Load(pID)

        Catch er As Exception
            MsgBox("Error in CEmail." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_ID = 0
            m_Ord_Guid = m_oOrder.Ordhead.Ord_GUID

            Call Reset()

        Catch er As Exception
            MsgBox("Error in CEmail." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Load is made private, it is loaded automatically on New.
    Private Sub Load(ByVal pintID As Integer)

        Try
            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String
            strSql = _
            "SELECT * " & _
            "FROM   OEI_Email WITH (Nolock) " & _
            "WHERE  ID = " & pintID & " "

            dt = db.DataTable(strSql)

            Call LoadLine(dt.Rows(0))

        Catch er As Exception
            MsgBox("Error in CEmail." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Function Load(ByVal pstrOrd_Guid As String) As DataTable

    '    Load = New DataTable

    '    Try
    '        Dim db As New cDBA
    '        Dim dt As DataTable

    '        Dim strSql As String

    '        strSql = _
    '        "SELECT     * " & _
    '        "FROM       OEI_Email H WITH (Nolock) " & _
    '        "WHERE      Ord_Guid = '" & pstrOrd_Guid & "' "

    '        dt = db.DataTable(strSql)
    '        m_OEI_Ord_No = dt.Rows(0).Item("Oei_Ord_No")

    '        strSql = _
    '        "SELECT     S.ID, S.Ord_Guid, S.Cmt, S.Line_Seq_No, H.OEI_Ord_No, S.Ord_No " & _
    '        "FROM       OEI_LinCmt S WITH (Nolock) " & _
    '        "INNER JOIN OEI_OrdHdr H WITH (Nolock) ON H.Ord_Guid = S.Ord_Guid " & _
    '        "WHERE      H.Ord_Guid = '" & pstrOrd_Guid & "' " & _
    '        "ORDER BY   Line_Seq_No "

    '        Load = db.DataTable(strSql)

    '    Catch er As Exception
    '        MsgBox("Error in CEmail." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Function

    ' Load is made private, it is loaded automatically on New.
    ' Loads the Comment fields from the DataRow record
    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            m_ID = pdrRow.Item("ID")
            m_Email_From = pdrRow.Item("Email_From").ToString
            m_Email_To = pdrRow.Item("Email_To").ToString
            m_Email_To_Name = pdrRow.Item("Email_To_Name").ToString
            m_Email_CC = pdrRow.Item("Email_CC").ToString
            m_Email_CC_Name = pdrRow.Item("Email_CC_Name").ToString
            m_Subject = pdrRow.Item("Subject").ToString
            m_Body = pdrRow.Item("Body").ToString
            m_Ord_Guid = pdrRow.Item("Ord_Guid").ToString
            m_Ord_No = pdrRow.Item("Ord_No").ToString
            If Not (pdrRow.Item("CreateTS").Equals(DBNull.Value)) Then m_CreateTS = pdrRow.Item("CreateTS").ToString
            If Not (pdrRow.Item("SendTS").Equals(DBNull.Value)) Then m_SendTS = pdrRow.Item("SendTS").ToString
            m_UserID = pdrRow.Item("UserID").ToString
            'If Not (pdrRow.Item("Line_Seq_No").Equals(DBNull.Value)) Then m_Line_Seq_No = pdrRow.Item("Line_Seq_No")

        Catch er As Exception
            MsgBox("Error in CEmail." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Saves the Comment fields into the DataRow record
    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            pdrRow.Item("Email_From") = m_Email_From
            pdrRow.Item("Email_To") = m_Email_To
            pdrRow.Item("Email_To_Name") = m_Email_To_Name
            pdrRow.Item("Email_CC") = m_Email_CC
            pdrRow.Item("Email_CC_Name") = m_Email_CC_Name
            pdrRow.Item("Subject") = m_Subject
            pdrRow.Item("Body") = m_Body
            pdrRow.Item("Ord_Guid") = m_Ord_Guid
            pdrRow.Item("CreateTS") = Now
            pdrRow.Item("UserID") = m_UserID
            ' We do not save SendTS and Ord_No because they are changed by the 
            ' trigger update and the Spector Emailer. We let it stay NULL.

        Catch er As Exception
            MsgBox("Error in CEmail." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    ' Deletes the current Comment from the database, if it exists.
    Public Sub Delete()

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            If m_ID = 0 Then Exit Sub

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            strSql = _
            "DELETE FROM OEI_Email " & _
            "WHERE  ID = " & m_ID & " AND " & _
            "       Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in CEmail." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Reset every non mandatory field of a Comment
    Public Sub Reset()

        ' Resets the locals to empty values, all but ID, Ord_No and Ord_Guid
        m_Email_From = g_User.Mail ' String.Empty
        m_Email_To = String.Empty
        m_Email_To_Name = String.Empty
        m_Email_CC = String.Empty
        m_Email_CC_Name = String.Empty
        m_Subject = String.Empty
        m_Body = String.Empty
        m_Ord_No = String.Empty
        m_CreateTS = Nothing
        m_SendTS = Nothing
        m_UserID = String.Empty

        'm_Line_Seq_No = 0

    End Sub

    ' Update the current Comment into the database, or creates it if not existing
    Public Sub Save()

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String

            strSql = _
            "SELECT * " & _
            "FROM   OEI_Email " & _
            "WHERE  ID = " & m_ID & " AND " & _
            "       Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' "

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

            If m_ID <> 0 Then m_ID = db.DataTable("SELECT MAX(ID) AS MaxID FROM OEI_Email WITH (Nolock)").Rows(0).Item("MaxID")

        Catch er As Exception

            MsgBox("Error in CEmail." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property Body() As String
        Get
            Body = m_Body
        End Get
        Set(ByVal value As String)
            m_Body = value
        End Set
    End Property
    Public Property CreateTS() As Date
        Get
            CreateTS = m_CreateTS
        End Get
        Set(ByVal value As Date)
            m_CreateTS = value
        End Set
    End Property
    Public Property Email_CC() As String
        Get
            Email_CC = m_Email_CC
        End Get
        Set(ByVal value As String)
            m_Email_CC = value
        End Set
    End Property
    Public Property Email_CC_Name() As String
        Get
            Email_CC_Name = m_Email_CC_Name
        End Get
        Set(ByVal value As String)
            m_Email_CC_Name = value
        End Set
    End Property
    Public Property Email_From() As String
        Get
            Email_From = m_Email_From
        End Get
        Set(ByVal value As String)
            m_Email_From = value
        End Set
    End Property
    Public Property Email_To() As String
        Get
            Email_To = m_Email_To
        End Get
        Set(ByVal value As String)
            m_Email_To = value
        End Set
    End Property
    Public Property Email_To_Name() As String
        Get
            Email_To_Name = m_Email_To_Name
        End Get
        Set(ByVal value As String)
            m_Email_To_Name = value
        End Set
    End Property
    Public Property ID() As Integer
        Get
            ID = m_ID
        End Get
        Set(ByVal value As Integer)
            m_ID = value
        End Set
    End Property
    Public Property Ord_Guid() As String
        Get
            Ord_Guid = m_Ord_Guid
        End Get
        Set(ByVal value As String)
            m_Ord_Guid = value
        End Set
    End Property
    Public Property Ord_No() As String
        Get
            Ord_No = m_Ord_No
        End Get
        Set(ByVal value As String)
            m_Ord_No = value
        End Set
    End Property
    Public Property SendTS() As Date
        Get
            SendTS = m_SendTS
        End Get
        Set(ByVal value As Date)
            m_SendTS = value
        End Set
    End Property
    Public Property Subject() As String
        Get
            Subject = m_Subject
        End Get
        Set(ByVal value As String)
            m_Subject = value
        End Set
    End Property
    Public Property UserID() As String
        Get
            UserID = m_UserID
        End Get
        Set(ByVal value As String)
            m_UserID = value
        End Set
    End Property

#End Region

End Class
