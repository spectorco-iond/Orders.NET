Public Class cMDB_Logo

    Private m_Logo_ID As Integer = 0
    Private m_Cus_No As String
    Private m_Logo_Name As String
    Private m_Logo_Doc As Image
    Private m_Logo_Doc_Array As Byte()
    Private m_File_Url As String
    'Private m_Line_Seq_No As Integer

    Private m_DataTable As DataTable
    Private oImage As New cImage()

#Region "Public constructors ##############################################"

    'Public Sub New()
    Public Sub New()

        Call Init()

    End Sub

    'Public Sub New(ByVal pID As Integer)
    Public Sub New(ByVal pLogo_ID As Integer)

        Try

            Call Init()

            m_Logo_ID = pLogo_ID

            Call Load(pLogo_ID)

        Catch er As Exception
            MsgBox("Error in cMDB_Logo." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ''Public Sub New(ByVal pstrComment_Name As String, ByVal pstrComment_Name_Value As String)
    'Public Sub New(ByVal pLogo_Doc As Image)

    '    Try

    '        Call Init()

    '        m_Logo_Doc = pLogo_Doc

    '        Call Load(pLogo_Doc)

    '    Catch er As Exception
    '        MsgBox("Error in cMDB_Logo." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_Logo_ID = 0
            m_Logo_Name = String.Empty
            m_Logo_Doc = Nothing

            Call Reset()

        Catch er As Exception
            MsgBox("Error in cMDB_Logo." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
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
            "FROM   MDB_Logo C WITH (Nolock) " & _
            "WHERE  Logo_ID = " & pintID & " "

            dt = db.DataTable(strSql)

            Call LoadLine(dt.Rows(0))

        Catch er As Exception
            MsgBox("Error in cMDB_Logo." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Function Load(ByVal pstrLogo_Doc As String) As DataTable

    '    Load = New DataTable

    '    Try
    '        Dim db As New cDBA
    '        Dim dt As DataTable

    '        Dim strSql As String

    '        strSql = _
    '        "SELECT     * " & _
    '        "FROM       OEI_OrdHdr H WITH (Nolock) " & _
    '        "WHERE      Logo_Doc = '" & pstrLogo_Doc & "' "

    '        dt = db.DataTable(strSql)
    '        m_Cus_No = dt.Rows(0).Item("Cus_No")

    '        strSql = _
    '        "SELECT     S.ID, S.Logo_Doc, S.File_Url, S.Line_Seq_No, H.Cus_No, S.Logo_Name " & _
    '        "FROM       OEI_LinFile_Url S WITH (Nolock) " & _
    '        "INNER JOIN OEI_OrdHdr H WITH (Nolock) ON H.Logo_Doc = S.Logo_Doc " & _
    '        "WHERE      H.Logo_Doc = '" & pstrLogo_Doc & "' " & _
    '        "ORDER BY   Line_Seq_No "

    '        Load = db.DataTable(strSql)

    '    Catch er As Exception
    '        MsgBox("Error in cMDB_Logo." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Function

    ' Load is made private, it is loaded automatically on New.
    ' Loads the Comment fields from the DataRow record
    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            If Not (pdrRow.Item("Logo_ID").Equals(DBNull.Value)) Then m_Logo_ID = pdrRow.Item("Logo_ID")
            If Not (pdrRow.Item("Logo_Name").Equals(DBNull.Value)) Then m_Logo_Name = pdrRow.Item("Logo_Name").ToString
            If Not (pdrRow.Item("Logo_Doc").Equals(DBNull.Value)) Then oImage.SetImage(pdrRow.Item("Logo_Doc"), m_Logo_Doc, m_Logo_Doc_Array)
            'If Not (pdrRow.Item("Logo_Doc").Equals(DBNull.Value)) Then oImage.SetImage(pdrRow.Item("Logo_Doc"), m_Logo_Doc)
            If Not (pdrRow.Item("File_Url").Equals(DBNull.Value)) Then m_File_Url = pdrRow.Item("File_Url").ToString
            'If Not (pdrRow.Item("Line_Seq_No").Equals(DBNull.Value)) Then m_Line_Seq_No = pdrRow.Item("Line_Seq_No")

        Catch er As Exception
            MsgBox("Error in cMDB_Logo." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Saves the Comment fields into the DataRow record
    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            pdrRow.Item("Cus_No") = m_Cus_No
            pdrRow.Item("Logo_Name") = m_Logo_Name
            Dim baImage As Byte() = Nothing
            oImage.SetImageAsByteArray(m_Logo_Doc, baImage)
            pdrRow.Item("Logo_Doc") = baImage
            pdrRow.Item("File_Url") = m_File_Url

        Catch er As Exception
            MsgBox("Error in cMDB_Logo." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    ' Deletes the current Comment from the database, if it exists.
    Public Sub Delete()

        Try
            If m_Logo_ID = 0 Then Exit Sub

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            strSql = _
            "DELETE FROM MDB_Logo " & _
            "WHERE  Logo_ID = " & m_Logo_ID & " "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in cMDB_Logo." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Reset every non mandatory field of a Comment
    Public Sub Reset()

        ' Resets the locals to empty values, all but ID, Logo_Name and Logo_Doc
        m_Logo_Name = String.Empty
        m_Logo_Doc = Nothing
        m_File_Url = String.Empty
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
            "FROM   MDB_Logo " & _
            "WHERE  Logo_ID = " & m_Logo_ID & " "

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
            MsgBox("Error in cMDB_Logo." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property File_Url() As String
        Get
            File_Url = m_File_Url
        End Get
        Set(ByVal value As String)
            m_File_Url = value
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
    Public Property Logo_ID() As Integer
        Get
            Logo_ID = m_Logo_ID
        End Get
        Set(ByVal value As Integer)
            m_Logo_ID = value
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
    Public Property Logo_Doc() As Image
        Get
            Logo_Doc = m_Logo_Doc
        End Get
        Set(ByVal value As Image)
            m_Logo_Doc = value
        End Set
    End Property
    Public Property Logo_Doc_Array() As Byte()
        Get
            Logo_Doc_Array = m_Logo_Doc_Array
        End Get
        Set(ByVal value As Byte())
            m_Logo_Doc_Array = value
        End Set
    End Property
    Public Property Cus_No() As String
        Get
            Cus_No = m_Cus_No
        End Get
        Set(ByVal value As String)
            m_Cus_No = value
        End Set
    End Property
    Public Property Logo_Name() As String
        Get
            Logo_Name = m_Logo_Name
        End Get
        Set(ByVal value As String)
            m_Logo_Name = value
        End Set
    End Property

#End Region


End Class
