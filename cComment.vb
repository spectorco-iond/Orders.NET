Public Class cComment

    Private m_ID As Integer = 0
    Private m_OEI_Ord_No As String
    Private m_Ord_No As String
    Private m_Ord_Guid As String
    Private m_Cmt As String
    Private m_Item_Cd As String
    'Private m_Line_Seq_No As Integer

    Private m_DataTable As DataTable

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
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Sub New(ByVal pstrComment_Name As String, ByVal pstrComment_Name_Value As String)
    Public Sub New(ByVal pstrOrd_Guid As String)

        Try

            Call Init()

            m_Ord_Guid = pstrOrd_Guid

            Call Load(pstrOrd_Guid)

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_ID = 0
            m_Ord_No = String.Empty
            m_Ord_Guid = String.Empty

            Call Reset()

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
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
            "FROM   OEI_LinCmt C WITH (Nolock) " & _
            "WHERE  ID = " & pintID & " "

            dt = db.DataTable(strSql)

            Call LoadLine(dt.Rows(0))

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function Load(ByVal pstrOrd_Guid As String) As DataTable

        Load = New DataTable

        Try
            Dim db As New cDBA
            Dim dt As DataTable

            Dim strSql As String

            strSql = _
            "SELECT     * " & _
            "FROM       OEI_OrdHdr H WITH (Nolock) " & _
            "WHERE      Ord_Guid = '" & pstrOrd_Guid & "' "

            dt = db.DataTable(strSql)
            m_OEI_Ord_No = dt.Rows(0).Item("Oei_Ord_No")

            strSql = _
            "SELECT     S.ID, S.Ord_Guid, S.Cmt, S.Line_Seq_No, H.OEI_Ord_No, S.Ord_No, S.Item_Cd " & _
            "FROM       OEI_LinCmt S WITH (Nolock) " & _
            "INNER JOIN OEI_OrdHdr H WITH (Nolock) ON H.Ord_Guid = S.Ord_Guid " & _
            "WHERE      H.Ord_Guid = '" & pstrOrd_Guid & "' " & _
            "ORDER BY   Line_Seq_No "

            Load = db.DataTable(strSql)

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Load is made private, it is loaded automatically on New.
    ' Loads the Comment fields from the DataRow record
    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            If Not (pdrRow.Item("ID").Equals(DBNull.Value)) Then m_ID = pdrRow.Item("ID")
            If Not (pdrRow.Item("Ord_No").Equals(DBNull.Value)) Then m_Ord_No = pdrRow.Item("Ord_No").ToString
            If Not (pdrRow.Item("Ord_Guid").Equals(DBNull.Value)) Then m_Ord_Guid = pdrRow.Item("Ord_Guid").ToString
            If Not (pdrRow.Item("Cmt").Equals(DBNull.Value)) Then m_Cmt = pdrRow.Item("Cmt").ToString
            If Not (pdrRow.Item("Item_Cd").Equals(DBNull.Value)) Then m_Item_Cd = pdrRow.Item("Item_Cd").ToString
            'If Not (pdrRow.Item("Line_Seq_No").Equals(DBNull.Value)) Then m_Line_Seq_No = pdrRow.Item("Line_Seq_No")

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Saves the Comment fields into the DataRow record
    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            pdrRow.Item("Ord_Guid") = m_Ord_Guid

            pdrRow.Item("Cmt") = m_Cmt ' m_Ord_No
            pdrRow.Item("Item_Cd") = m_Item_Cd ' m_Ord_No

            If pdrRow.Item("Line_Seq_No").Equals(DBNull.Value) Then
                Dim dt As DataTable
                Dim db As New cDBA
                dt = db.DataTable("SELECT MAX(ID) AS Max_ID FROM OEI_LinCmt WITH (Nolock) ")
                pdrRow.Item("Line_Seq_No") = dt.Rows(0).Item("Max_ID")
            End If

            'pdrRow.Item("Line_Seq_No") = m_Line_Seq_No

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
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
            "DELETE FROM OEI_LinCmt " & _
            "WHERE  ID = " & m_ID & " AND " & _
            "       Ord_Guid = '" & m_oOrder.Ordhead.Ord_GUID & "' "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Reset every non mandatory field of a Comment
    Public Sub Reset()

        ' Resets the locals to empty values, all but ID, Ord_No and Ord_Guid
        m_Ord_No = String.Empty
        m_Ord_Guid = String.Empty
        m_Cmt = String.Empty
        m_Item_Cd = String.Empty
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
            "FROM   OEI_LinCmt " & _
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

        Catch er As Exception
            MsgBox("Error in CComment." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property Cmt() As String
        Get
            Cmt = m_Cmt
        End Get
        Set(ByVal value As String)
            m_Cmt = value
        End Set
    End Property
    Public Property Comments() As DataTable
        Get
            Comments = m_DataTable
        End Get
        Set(ByVal value As DataTable)
            m_DataTable = value
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
    Public Property Item_Cd() As String
        Get
            Item_Cd = m_Item_Cd
        End Get
        Set(ByVal value As String)
            m_Item_Cd = value
        End Set
    End Property
    'Public Property Line_Seq_No() As Integer
    '    Get
    '        Line_Seq_No = m_Line_Seq_No
    '    End Get
    '    Set(ByVal value As Integer)
    '        m_Line_Seq_No = value
    '    End Set
    'End Property
    Public Property Ord_Guid() As String
        Get
            Ord_Guid = m_Ord_Guid
        End Get
        Set(ByVal value As String)
            m_Ord_Guid = value
        End Set
    End Property
    Public Property OEI_Ord_No() As String
        Get
            OEI_Ord_No = m_OEI_Ord_No
        End Get
        Set(ByVal value As String)
            m_OEI_Ord_No = value
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

#End Region

End Class
