Public Class cNote

    Private m_ID As Integer = 0
    Private m_Note_Name As String
    Private m_Note_Name_Value As String
    Private m_Note_Dt As Date
    Private m_Note_Tm As Date
    Private m_User_Name As String
    Private m_Note_Topic As String
    Private m_Notes As String
    Private m_Document_Field_1 As Integer ' Can be busted, SQL says decimal 12,0
    Private m_Filler_001 As String
    
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
            MsgBox("Error in CNote." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Sub New(ByVal pstrNote_Name As String, ByVal pstrNote_Name_Value As String)
    Public Sub New(ByVal pstrNote_Name As String, ByVal pstrNote_Name_Value As String)

        Try

            Call Init()

            m_Note_Name = pstrNote_Name
            m_Note_Name_Value = pstrNote_Name_Value

            Call Load(pstrNote_Name, pstrNote_Name_Value)

        Catch er As Exception
            MsgBox("Error in CNote." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_ID = 0
            m_Note_Name = String.Empty
            m_Note_Name_Value = String.Empty

            Call Reset()

        Catch er As Exception
            MsgBox("Error in CNote." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
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
            "FROM   SYSNOTES_SQL S WITH (Nolock) " & _
            "WHERE  ID = " & pintID & " "

            dt = db.DataTable(strSql)

            Call LoadLine(dt.Rows(0))

        Catch er As Exception
            MsgBox("Error in CNote." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Load is made private, it is loaded automatically on New.
    Private Sub Load(ByVal pstrNote_Name As String, ByVal pstrNote_Name_Value As String)

        Try
            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            Select Case pstrNote_Name
                Case "OE-ORDER-NO"
                    strSql = _
                    "SELECT * " & _
                    "FROM   SYSNOTES_SQL S WITH (Nolock) " & _
                    "WHERE  Note_Name = '" & pstrNote_Name & "' AND " & _
                    "       Filler_001 = '" & m_oOrder.Ordhead.Ord_GUID & "' "

                Case "IM-ITEM"
                    strSql = _
                    "SELECT * " & _
                    "FROM   SYSNOTES_SQL S WITH (Nolock) " & _
                    "WHERE  Note_Name = '" & pstrNote_Name & "' AND " & _
                    "       Filler_001 = '" & g_oOrdline.Item_Guid & "' "

                Case Else
                    strSql = _
                    "SELECT * " & _
                    "FROM   SYSNOTES_SQL S WITH (Nolock) " & _
                    "WHERE  Note_Name = '" & pstrNote_Name & "' AND " & _
                    "       Note_Name_Value = '" & pstrNote_Name_Value & "' "

            End Select

            dt = db.DataTable(strSql)

            Call LoadLine(dt.Rows(0))

        Catch er As Exception
            MsgBox("Error in CNote." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Loads the note fields from the DataRow record
    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            If Not (pdrRow.Item("Note_Name").Equals(DBNull.Value)) Then m_Note_Name = pdrRow.Item("Note_Name").ToString
            If Not (pdrRow.Item("Note_Name_Value").Equals(DBNull.Value)) Then m_Note_Name_Value = pdrRow.Item("Note_Name_Value").ToString
            If Not (pdrRow.Item("Note_Dt").Equals(DBNull.Value)) Then m_Note_Dt = pdrRow.Item("Note_Dt")
            If Not (pdrRow.Item("Note_Tm").Equals(DBNull.Value)) Then m_Note_Tm = pdrRow.Item("Note_Tm")
            If Not (pdrRow.Item("User_Name").Equals(DBNull.Value)) Then m_User_Name = pdrRow.Item("User_Name").ToString
            If Not (pdrRow.Item("Note_Topic").Equals(DBNull.Value)) Then m_Note_Topic = pdrRow.Item("Note_Topic").ToString
            If Not (pdrRow.Item("Notes").Equals(DBNull.Value)) Then m_Notes = pdrRow.Item("Notes").ToString
            If Not (pdrRow.Item("Document_Field_1").Equals(DBNull.Value)) Then m_Document_Field_1 = pdrRow.Item("Document_Field_1")
            If Not (pdrRow.Item("Filler_001").Equals(DBNull.Value)) Then m_Filler_001 = pdrRow.Item("Filler_001").ToString
            If Not (pdrRow.Item("ID").Equals(DBNull.Value)) Then m_ID = pdrRow.Item("ID")

        Catch er As Exception
            MsgBox("Error in CNote." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Saves the note fields into the DataRow record
    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            pdrRow.Item("Note_Name") = m_Note_Name
            pdrRow.Item("Note_Name_Value") = m_Note_Name_Value
            If m_ID = 0 Then
                pdrRow.Item("Note_Dt") = Date.Now.Date
                pdrRow.Item("Note_Tm") = Date.Now
            Else
                pdrRow.Item("Note_Dt") = m_Note_Dt.Date
                pdrRow.Item("Note_Tm") = m_Note_Tm
            End If

            pdrRow.Item("User_Name") = Trim(g_User.Lock_Code) ' m_User_Name
            pdrRow.Item("Note_Topic") = m_Note_Topic
            pdrRow.Item("Notes") = m_Notes
            pdrRow.Item("Document_Field_1") = m_Document_Field_1
            If m_Note_Name = "OE-ORDER-NO" Then
                pdrRow.Item("Filler_001") = m_oOrder.Ordhead.Ord_GUID
            End If
            If m_Note_Name = "IM-ITEM" Then ' IM-ITEM ' OE-ITEM-LOC POUR LA RECHERCHE
                pdrRow.Item("Filler_001") = g_oOrdline.Item_Guid
            End If
            ' Let Filler_001 NULL
            'pdrRow.Item("Filler_001") = m_Filler_001
            ' Don't resave ID, it saves as is or create a new one.
            'pdrRow.Item("ID") = m_ID

        Catch er As Exception
            MsgBox("Error in CNote." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    ' Deletes the current note from the database, if it exists.
    Public Sub Delete()

        If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

        Try

            If m_ID = 0 Then Exit Sub

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            Select Case m_Note_Name
                Case "OE-ORDER-NO"
                    strSql = _
                    "DELETE FROM SYSNOTES_SQL " & _
                    "WHERE  ID = " & m_ID & " AND " & _
                    "       Note_Name = '" & SqlCompliantString(m_Note_Name) & "' AND " & _
                    "       Filler_001 = '" & m_oOrder.Ordhead.Ord_GUID & "' "

                Case "IM-ITEM"
                    strSql = _
                    "DELETE FROM SYSNOTES_SQL " & _
                    "WHERE  ID = " & m_ID & " AND " & _
                    "       Note_Name = '" & SqlCompliantString(m_Note_Name) & "' AND " & _
                    "       Filler_001 = '" & g_oOrdline.Item_Guid & "' "

                Case Else
                    strSql = _
                    "DELETE FROM SYSNOTES_SQL " & _
                    "WHERE  ID = " & m_ID & " AND " & _
                    "       Note_Name = '" & SqlCompliantString(m_Note_Name) & "' AND " & _
                    "       Note_Name_Value = '" & m_Note_Name_Value & "' "

            End Select

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in CNote." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Reset every non mandatory field of a note
    Public Sub Reset()

        ' Resets the locals to empty values, all but ID, Note_Name and Note_Name_Value
        m_Note_Dt = New Date
        m_Note_Tm = New Date
        m_User_Name = g_User.Lock_Code ' String.Empty
        m_Note_Topic = String.Empty
        m_Notes = String.Empty
        m_Document_Field_1 = 0 ' Can be busted, SQL says decimal 12,0
        m_Filler_001 = String.Empty

    End Sub

    ' Update the current note into the database, or creates it if not existing
    Public Sub Save()

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String

            Select Case m_Note_Name
                Case "OE-ORDER-NO"
                    strSql = _
                    "SELECT * " & _
                    "FROM   SYSNOTES_SQL S " & _
                    "WHERE  ID = " & m_ID & " AND " & _
                    "       Note_Name = '" & m_Note_Name & "' AND " & _
                    "       Filler_001 = '" & m_oOrder.Ordhead.Ord_GUID & "' "

                Case "IM-ITEM"
                    strSql = _
                    "SELECT * " & _
                    "FROM   SYSNOTES_SQL S " & _
                    "WHERE  ID = " & m_ID & " AND " & _
                    "       Note_Name = '" & m_Note_Name & "' AND " & _
                    "       Filler_001 = '" & g_oOrdline.Item_Guid & "' "

                Case Else
                    strSql = _
                    "SELECT * " & _
                    "FROM   SYSNOTES_SQL S " & _
                    "WHERE  ID = " & m_ID & " AND " & _
                    "       Note_Name = '" & m_Note_Name & "' AND " & _
                    "       Note_Name_Value = '" & m_Note_Name_Value & "' "

            End Select

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
            MsgBox("Error in CNote." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property Note_Name() As String
        Get
            Note_Name = m_Note_Name
        End Get
        Set(ByVal value As String)
            m_Note_Name = value
        End Set
    End Property
    Public Property Note_Name_Value() As String
        Get
            Note_Name_Value = m_Note_Name_Value
        End Get
        Set(ByVal value As String)
            m_Note_Name_Value = value
        End Set
    End Property
    Public Property Note_Dt() As Date
        Get
            Note_Dt = m_Note_Dt
        End Get
        Set(ByVal value As Date)
            m_Note_Dt = value
        End Set
    End Property
    Public Property Note_Tm() As Date
        Get
            Note_Tm = m_Note_Tm
        End Get
        Set(ByVal value As Date)
            m_Note_Tm = value
        End Set
    End Property
    Public Property User_Name() As String
        Get
            User_Name = m_User_Name
        End Get
        Set(ByVal value As String)
            m_User_Name = value
        End Set
    End Property
    Public Property Note_Topic() As String
        Get
            Note_Topic = m_Note_Topic
        End Get
        Set(ByVal value As String)
            m_Note_Topic = value
        End Set
    End Property
    Public Property Notes() As String
        Get
            Notes = m_Notes
        End Get
        Set(ByVal value As String)
            m_Notes = value
        End Set
    End Property
    Public Property Document_Field_1() As Integer
        Get
            Document_Field_1 = m_Document_Field_1
        End Get
        Set(ByVal value As Integer)
            m_Document_Field_1 = value
        End Set
    End Property
    Public Property Filler_001() As String
        Get
            Filler_001 = m_Filler_001
        End Get
        Set(ByVal value As String)
            m_Filler_001 = value
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

#End Region

End Class
