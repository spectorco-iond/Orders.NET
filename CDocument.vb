Imports System.IO
Imports System.Text

Public Class CDocument

    ' Sql definitions in CreateSql sub 
    'Private m_Sql(6) As String

    Private m_DocID As Integer
    Private m_Ord_No As String
    Private m_HeaderID As Integer
    Private m_DocType As String
    Private m_DocTypeID As Integer
    Private m_DocDescription As String
    Private m_DocFile As FileInfo
    Private m_DocName As String
    Private m_Document As Object
    Private m_Document_Array As Byte()
    Private m_DocumentText As String = String.Empty
    Private m_CreateDate As Date
    Private m_ArtworkPurged As Integer
    Private m_Ord_Guid As String = String.Empty
    Private m_Item_Guid As String = String.Empty

    Private m_AutoSave As Boolean = False

    Private m_Email As Stream = Nothing
    Private m_Contents As MemoryStream = Nothing
    Private m_DeleteStatus As Integer = 0

    'Private m_DocFile As String
    Private m_OnDrive As Integer ' tinyint 0-255
    Private m_Added_By_User As String

    Public Sub New()

        Call Init()

        'Call CreateSql()

    End Sub

    '' when moving a document, as it is already saved, use this signature.
    'Public Sub New(ByVal pstrOrd_Guid As String, ByVal pstrItem_Guid As String, ByVal pstrFilePath As String)

    '    m_Ord_Guid = pstrOrd_Guid
    '    m_Item_Guid = pstrItem_Guid

    '    Call Reset()

    '    'Call CreateSql()

    'End Sub

    ' To create a new document, use this signature.
    Public Sub New(ByVal pstrOrd_Guid As String, ByVal pstrItem_Guid As String, ByVal pstrFilePath As String, ByVal pintDocTypeID As Integer, ByVal pstrDocType As String)

        Try
            m_Ord_Guid = pstrOrd_Guid
            m_Item_Guid = pstrItem_Guid

            Call Reset()

            m_DocTypeID = pintDocTypeID
            m_DocType = pstrDocType

            DocFile = pstrFilePath

            'Call CreateSql()

            m_AutoSave = True

            Call Save()

        Catch er As Exception
            MsgBox("Error in cDocument." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' To create a new document, use this signature.
    Public Sub New(ByVal pstrOrd_Guid As String, ByVal pstrItem_Guid As String, ByVal pintDocTypeID As Integer, ByVal pstrDocType As String, ByVal poEmail As Object, ByVal poContents As Object)

        Try
            m_Ord_Guid = pstrOrd_Guid
            m_Item_Guid = pstrItem_Guid

            Call Reset()

            m_DocTypeID = pintDocTypeID
            m_DocType = pstrDocType

            m_Contents = DirectCast(poContents, MemoryStream)
            Email = DirectCast(poEmail, Stream)
            'DocFile = pstrFilePath

            'Call CreateSql()

            m_AutoSave = True

            Call Save()

        Catch er As Exception
            MsgBox("Error in cDocument." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Private Sub CreateSql()

    '    Try
    '        'If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

    '        ''"FROM       OEI_Documents WITH (Nolock) " & _
    '        'm_Sql(SqlMode.SelectNoLock) = _
    '        '"SELECT     * " & _
    '        '"FROM       Exact_Traveler_Document WITH (Nolock) " & _
    '        '"WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND Item_Guid = '" & m_Item_Guid & "' AND DocName = '" & SqlCompliantString(m_DocName) & "' AND DocTypeID = " & m_DocTypeID

    '        ''"FROM       OEI_Documents " & _
    '        'm_Sql(SqlMode.SelectUpdLock) = _
    '        '"SELECT     * " & _
    '        '"FROM       Exact_Traveler_Document " & _
    '        '"WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND Item_Guid = '" & m_Item_Guid & "' AND DocName = '" & SqlCompliantString(m_DocName) & "' AND DocTypeID = " & m_DocTypeID

    '        ''"DELETE     FROM OEI_Documents " & _
    '        'm_Sql(SqlMode.Delete) = _
    '        '"DELETE     FROM Exact_Traveler_Document " & _
    '        '"WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND Item_Guid = '" & m_Item_Guid & "' AND DocName = '" & SqlCompliantString(m_DocName) & "' "


    '    Catch oe_er As OEException
    '        MsgBox(oe_er.Message)
    '    Catch er As Exception
    '        MsgBox("Error in cDocument." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Public Sub Delete()

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            'Call CreateSql()
            'If Trim(m_DocName).Length = 0 Then Throw New OEException(OEError.Nothing_To_Delete)

            'If m_DocID <> 0 Then Exit Sub 'Throw New OEException(OEError.Exit_Sub, False, False)

            Dim oMsgBox As New MsgBoxResult
            oMsgBox = MsgBox("Do you want to remove file " & m_DocName & " ?", MsgBoxStyle.YesNoCancel, "Remove document")
            If oMsgBox = MsgBoxResult.Yes Then

                Dim conn As New cDBA
                Dim strSql As String = "" & _
                "DELETE     FROM Exact_Traveler_Document " & _
                "WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND " & _
                "           DocName = '" & SqlCompliantString(m_DocName) & "' AND " & _
                "           DocTypeID = '" & SqlCompliantString(m_DocTypeID) & "' AND " & _
                "           DocID = " & m_DocID

                'conn.Execute(m_Sql(SqlMode.Delete))
                conn.Execute(strSql)

                m_DeleteStatus = mGlobal.DeleteStatus.Deleted

            ElseIf oMsgBox = MsgBoxResult.No Then
                Exit Sub 'Throw New OEException(OEError.Exit_Sub, False, False)
            Else
                Exit Sub 'Throw New OEException(OEError.Exit_Sub, False, False)
            End If

        Catch oe_er As OEException
            Throw oe_er
        Catch er As Exception
            MsgBox("Error in CDocument." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Init()

        m_Ord_Guid = String.Empty
        m_Item_Guid = String.Empty

        Call Reset()

    End Sub

    Public Sub Load()

        Dim conn As New cDBA
        Dim dt As DataTable
        Dim strSql As String

        Try

            strSql = _
            "SELECT     * " & _
            "FROM       Exact_Traveler_Document WITH (Nolock) " & _
            "WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND Item_Guid = '" & m_Item_Guid & "' AND DocName = '" & SqlCompliantString(m_DocName) & "' AND DocTypeID = " & m_DocTypeID

            dt = conn.DataTable(strSql) ' m_Sql(SqlMode.SelectNoLock))
            If dt.Rows.Count <> 0 Then

                m_DocID = dt.Rows(0).Item("DocID").ToString
                m_Ord_No = IIf(dt.Rows(0).Item("Ord_No").Equals(DBNull.Value), "", Trim(dt.Rows(0).Item("Ord_No")))
                m_HeaderID = IIf(dt.Rows(0).Item("HeaderID").Equals(DBNull.Value), 0, dt.Rows(0).Item("HeaderID"))
                m_DocType = IIf(dt.Rows(0).Item("DocType").Equals(DBNull.Value), "", Trim(dt.Rows(0).Item("DocType")))
                m_DocTypeID = IIf(dt.Rows(0).Item("DocTypeID").Equals(DBNull.Value), 0, dt.Rows(0).Item("DocTypeID"))
                m_DocDescription = IIf(dt.Rows(0).Item("DocDescription").Equals(DBNull.Value), "", Trim(dt.Rows(0).Item("DocDescription")))
                m_DocName = IIf(dt.Rows(0).Item("DocName").Equals(DBNull.Value), "", Trim(dt.Rows(0).Item("DocName")))
                m_Document = IIf(dt.Rows(0).Item("Document").Equals(DBNull.Value), "", Trim(dt.Rows(0).Item("Document")))
                m_CreateDate = IIf(dt.Rows(0).Item("CreateDate").Equals(DBNull.Value), "", dt.Rows(0).Item("CreateDate"))
                m_ArtworkPurged = IIf(dt.Rows(0).Item("ArtworkPurged").Equals(DBNull.Value), 0, dt.Rows(0).Item("ArtworkPurged"))
                m_Ord_Guid = IIf(dt.Rows(0).Item("Ord_Guid").Equals(DBNull.Value), "", Trim(dt.Rows(0).Item("Ord_Guid")))
                m_Item_Guid = IIf(dt.Rows(0).Item("Item_Guid").Equals(DBNull.Value), "", Trim(dt.Rows(0).Item("Item_Guid")))
                m_OnDrive = IIf(dt.Rows(0).Item("OnDrive").Equals(DBNull.Value), "", Trim(dt.Rows(0).Item("OnDrive")))
                m_Added_By_User = IIf(dt.Rows(0).Item("Added_By_User").Equals(DBNull.Value), "", Trim(dt.Rows(0).Item("Added_By_User")))

            End If
            dt.Dispose()

        Catch er As Exception
            MsgBox("Error in CDocument." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Load(ByRef pRow As DataGridViewRow)

        Dim strValue As String

        Try
            For Each pRowCell As DataGridViewCell In pRow.Cells

                If pRowCell.Value Is Nothing Then
                    strValue = ""
                Else
                    strValue = pRowCell.Value.ToString
                End If

                Select Case pRowCell.DataGridView.Columns(pRowCell.ColumnIndex).Name

                    Case "DocID"
                        If strValue = "" Then strValue = "0"
                        m_DocID = CInt(strValue)

                    Case "Ord_No"
                        m_Ord_No = strValue

                    Case "HeaderID"
                        If strValue = "" Then strValue = "0"
                        m_HeaderID = CInt(strValue)

                    Case "DocType"
                        m_DocType = strValue

                    Case "DocTypeID"
                        If strValue = "" Then strValue = "0"
                        m_DocTypeID = CInt(strValue)

                    Case "DocDescription"
                        m_DocDescription = strValue

                    Case "DocFile"
                        If m_DocID = 0 Then
                            m_DocFile = New FileInfo(strValue)
                        End If

                    Case "DocName"
                        m_DocName = strValue

                        'Case "m_Document"
                        '    m_Document = strValue

                    Case "CreateDate"
                        m_CreateDate = strValue

                    Case "ArtworkPurged"
                        If strValue = "" Then strValue = "0"
                        m_ArtworkPurged = CInt(strValue)

                    Case "Ord_Guid"
                        m_Ord_Guid = strValue

                    Case "Item_Guid"
                        m_Item_Guid = strValue

                    Case "OnDrive"
                        m_OnDrive = CInt(strValue)

                    Case "Added_By_User"
                        m_Added_By_User = strValue

                    Case Else
                        ' Do nothing

                End Select

            Next

            'If m_DocID = 0 Then Call CreateSql()
            'Call CreateSql()

        Catch er As Exception
            MsgBox("Error in CDocument." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Reset()

        m_DocID = 0
        m_Ord_No = String.Empty
        m_HeaderID = 0
        m_DocType = String.Empty
        m_DocTypeID = 0
        m_DocDescription = String.Empty
        m_DocFile = Nothing
        m_DocName = String.Empty
        m_Document = Nothing ' New Object()
        m_DocumentText = String.Empty
        m_CreateDate = Now ' String.Empty
        m_ArtworkPurged = 0
        m_OnDrive = 0
        m_Added_By_User = String.Empty
        m_Email = Nothing

    End Sub

    Public Sub Save()

        Dim strSql As String

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Try

            'Call CreateSql()
            'If InStr(m_Sql(SqlMode.SelectUpdLock), "Ord_Guid = ''") Then Call CreateSql()

            'm_DocName = m_DocFile.Name

            Dim dt As New DataTable
            Dim conn As New cDBA
            strSql = _
                        "SELECT     * " & _
            "FROM       Exact_Traveler_Document " & _
            "WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND Item_Guid = '" & m_Item_Guid & "' AND DocName = '" & SqlCompliantString(m_DocName) & "' AND DocTypeID = " & m_DocTypeID

            dt = conn.DataTable(strSql) ' m_Sql(SqlMode.SelectUpdLock))

            Dim drDataRow As DataRow '= dsDataSet.Tables("OEI_ORDLIN").Rows(1)

            If dt.Rows.Count <> 0 Then
                drDataRow = dt.Rows(0)
                'conn.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            Else
                drDataRow = dt.NewRow()
                'conn.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
            End If

            If m_DocID <> 0 Then m_AutoSave = True

            Call SaveDataRow(drDataRow)

            If m_DocName = "" Or m_Ord_Guid = "" Then Exit Sub

            conn.DBDataTable.Rows.Add(drDataRow)

            If dt.Rows.Count = 0 Then
                conn.DBDataTable.Rows.Add(drDataRow)
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(conn.DBDataAdapter)
                conn.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
            Else
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(conn.DBDataAdapter)
                conn.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            End If

            conn.DBDataAdapter.Update(conn.DBDataTable)

            If m_DocID = 0 Then
                'Dim strSql As String = _
                strSql = _
                "SELECT MAX(DocID) AS DocID " & _
                "FROM   Exact_Traveler_Document WITH (Nolock) " & _
                "WHERE  Ord_Guid = '" & m_Ord_Guid & "' AND DocTypeID = " & m_DocTypeID & " AND DocName = '" & SqlCompliantString(m_DocName) & "' "
                dt = conn.DataTable(strSql)
                If dt.Rows.Count <> 0 Then
                    m_DocID = dt.Rows(0).Item("DocID")
                End If
            End If

        Catch er As Exception
            MsgBox("Error in CDocument." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Save(ByVal pintDocID As Integer)

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Try

            'Call CreateSql()
            'If InStr(m_Sql(SqlMode.SelectUpdLock), "Ord_Guid = ''") Then Call CreateSql()

            'm_DocName = m_DocFile.Name

            Dim dt As New DataTable
            Dim conn As New cDBA

            Dim strSql As String = _
            "SELECT     * " & _
            "FROM       Exact_Traveler_Document WITH (Nolock) " & _
            "WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND DocID = " & m_DocID

            dt = conn.DataTable(strSql)

            Dim drDataRow As DataRow '= dsDataSet.Tables("OEI_ORDLIN").Rows(1)

            If dt.Rows.Count <> 0 Then
                drDataRow = dt.Rows(0)
            Else
                drDataRow = dt.NewRow()
            End If

            If m_DocID <> 0 Then m_AutoSave = True

            Call SaveDataRow(drDataRow)

            If m_DocName = "" Or m_Ord_Guid = "" Then Exit Sub

            If dt.Rows.Count = 0 Then
                conn.DBDataTable.Rows.Add(drDataRow)
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(conn.DBDataAdapter)
                conn.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
            Else
                Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(conn.DBDataAdapter)
                conn.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            End If

            conn.DBDataAdapter.Update(conn.DBDataTable)

        Catch er As Exception
            MsgBox("Error in CDocument." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SaveDataRow(ByRef pRow As DataRow)

        Try

            pRow.Item("Ord_Guid") = m_Ord_Guid
            pRow.Item("Item_Guid") = m_Item_Guid

            pRow.Item("DocType") = m_DocType
            pRow.Item("DocTypeID") = m_DocTypeID
            pRow.Item("DocDescription") = m_DocDescription

            ' AutoSave is triggered by document load
            If m_AutoSave Then

                'pRow.Item("DocID") = m_DocID ' Identity, not needed

                pRow.Item("Ord_No") = m_Ord_No

                Dim mstream As New ADODB.Stream

                ' Dim mStream As New ADODB.Stream
                If m_DocID <> 0 Then
                    pRow.Item("DocID") = 0
                    pRow.Item("HeaderID") = 0
                    pRow.Item("DocFile") = m_DocName
                    pRow.Item("DocName") = m_DocName
                    pRow.Item("DocDescription") = m_DocDescription
                    If m_OnDrive = 2 Then
                        pRow.Item("Document") = m_DocumentText
                    Else
                        pRow.Item("Document") = GetDocumentRow.Item("Document")
                    End If
                Else

                    pRow.Item("HeaderID") = m_HeaderID
                    'pRow.Item("DocFile") = m_DocFile.FullName
                    'pRow.Item("DocName") = m_DocFile.Name
                    'pRow.Item("DocDescription") = Mid(m_DocFile.Name, 1, m_DocFile.Name.Length - m_DocFile.Extension.Length)

                    'Call SetSubjectForMSGFile(pRow)

                    'If m_Email Is Nothing Then

                    'If m_DocName <> "" Then
                    m_DocName = Trim(m_oOrder.Ordhead.OEI_Ord_No) & " - " & m_DocFile.Name
                    pRow.Item("DocName") = m_DocName ' Trim(m_oOrder.Ordhead.OEI_Ord_No) & " - " & m_DocFile.Name
                    pRow.Item("DocFile") = m_DocFile.FullName.Replace(m_DocFile.Name, Trim(m_oOrder.Ordhead.OEI_Ord_No) & " - " & m_DocFile.Name)
                    m_DocDescription = Trim(m_oOrder.Ordhead.OEI_Ord_No) & " - " & Mid(m_DocFile.Name, 1, m_DocFile.Name.Length - m_DocFile.Extension.Length)
                    pRow.Item("DocDescription") = m_DocDescription ' Trim(m_oOrder.Ordhead.OEI_Ord_No) & " - " & Mid(m_DocFile.Name, 1, m_DocFile.Name.Length - m_DocFile.Extension.Length)

                    If m_OnDrive = 2 Then
                        pRow.Item("Document") = m_DocumentText
                    Else
                        If m_Document Is Nothing Then
                            mstream.Type = ADODB.StreamTypeEnum.adTypeBinary
                            mstream.Open()
                            mstream.LoadFromFile(m_DocFile.FullName)
                            pRow.Item("Document") = mstream.Read
                        Else
                            If pRow.Item("Document").Equals(DBNull.Value) Then
                                If m_Document_Array Is Nothing Then
                                    Dim oimage As cImage = New cImage()
                                    Dim baImage As Byte() = Nothing
                                    oimage.SetImageAsByteArray(m_Document, baImage)
                                    pRow.Item("Document") = baImage
                                Else
                                    pRow.Item("Document") = m_Document_Array
                                End If
                                m_CreateDate = Now()
                                m_ArtworkPurged = False
                            End If
                        End If
                    End If
                    'End If

                    'Else
                    'pRow.Item("DocFile") = m_DocFile.FullName
                    'pRow.Item("DocName") = m_DocFile.Name
                    'pRow.Item("DocDescription") = m_DocDescription
                    'mstream.Type = ADODB.StreamTypeEnum.adTypeBinary
                    'mstream.Open()
                    ''For iPos As Integer = 1 To m_Email.Length
                    'mstream.Write(m_Email)
                    ''Next
                    ''mstream.Read()
                    'm_Email.Read(pRow.Item("Document"), 0, m_Email.Length - 1)
                    ''pRow.Item("Document") = m_Email.Read(
                    'End If

                End If

                    If pRow.Item("CreateDate").Equals(DBNull.Value) Then pRow.Item("CreateDate") = m_CreateDate
                    pRow.Item("ArtworkPurged") = m_ArtworkPurged

                    pRow.Item("OnDrive") = m_OnDrive
                    pRow.Item("Added_By_User") = m_Added_By_User

                    'Else
                    '    pRow.Item("DocName") = m_DocName
            End If

        Catch er As Exception
            MsgBox("Error in COrdline." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub UpdateDataGridViewRow(ByRef pRow As DataGridViewRow)

    End Sub

    Private Sub SetSubjectForMSGFile(ByRef pRow As DataRow)

        Try
            If m_DocFile.Extension.ToUpper = "MSG" Then


            End If

        Catch er As Exception
            ' In case Outlook is not installed, we're gonna do nothing here.
        End Try

    End Sub


    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////       VERIFY DOC ADD             /////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ' Bankers wants to ensure there is only ever ONE Final Artwork.
    ' So if they try to add another one, must give them the option of cancelling the add,
    ' or changing the previously "Final Artwork" to "Artwork"
    Private Function VerifyDocAdd() As Boolean ' ByRef NewHeaderID As Object, ByRef OrderNo As Object, ByRef DocType As Object, ByRef DocName As Object) As Boolean
        '    Private Function VerifyDocAdd(ByRef NewHeaderID As Object, ByRef OrderNo As Object, ByRef DocType As Object, ByRef DocName As Object) As Boolean

        ' This functionnality will not be implemented for now, as the Final Artworks will
        ' Not be entered by them.
        VerifyDocAdd = False

    End Function

    Public Function DocTypeDataTable() As DataTable

        DocTypeDataTable = New DataTable
        Dim db As New cDBA

        Dim strSql As String = _
        "SELECT		0 AS DocTypeID, '* All Types *' AS DocType " & _
        "UNION      " & _
        "SELECT     ISNULL(DocTypeID, 0) AS DocTypeID, " & _
        "           ISNULL(DocType, '') AS DocType " & _
        "FROM       Exact_Traveler_xref_Doctype WITH (Nolock) " & _
        "WHERE      Active = 1 AND DocType IS NOT NULL AND Use_In_OEI = 1 " & _
        "ORDER BY   DocType Asc "

        DocTypeDataTable = db.DataTable(strSql)

    End Function

    Public Property Added_By_User() As String
        Get
            Added_By_User = m_Added_By_User
        End Get
        Set(ByVal value As String)
            m_Added_By_User = value
        End Set
    End Property
    Public Property ArtworkPurged() As Integer
        Get
            ArtworkPurged = m_ArtworkPurged
        End Get
        Set(ByVal value As Integer)
            m_ArtworkPurged = value
        End Set
    End Property
    Public Property AutoSave() As Boolean
        Get
            AutoSave = m_AutoSave
        End Get
        Set(ByVal value As Boolean)
            m_AutoSave = value
        End Set
    End Property
    Public Property CreateDate() As Date
        Get
            CreateDate = m_CreateDate
        End Get
        Set(ByVal value As Date)
            m_CreateDate = value
        End Set
    End Property
    Public Property DeleteStatus() As Integer
        Get
            DeleteStatus = m_DeleteStatus
        End Get
        Set(ByVal value As Integer)
            m_DeleteStatus = value
        End Set
    End Property

    Public Property DocDescription() As String
        Get
            DocDescription = m_DocDescription
        End Get
        Set(ByVal value As String)
            m_DocDescription = value
        End Set
    End Property
    Public Property DocFile() As String
        Get
            DocFile = m_DocFile.Name
        End Get
        Set(ByVal value As String)
            If value = "" Then
                m_DocFile = Nothing
                m_DocName = ""
            Else
                m_DocFile = New FileInfo(value)
                m_DocName = m_DocFile.Name
            End If
        End Set
    End Property
    Public Property DocID() As Integer
        Get
            DocID = m_DocID
        End Get
        Set(ByVal value As Integer)
            m_DocID = value
        End Set
    End Property
    Public Property DocType() As String
        Get
            DocType = m_DocType
        End Get
        Set(ByVal value As String)
            m_DocType = value
        End Set
    End Property
    Public Property DocTypeID() As Integer
        Get
            DocTypeID = m_DocTypeID
        End Get
        Set(ByVal value As Integer)
            m_DocTypeID = value
        End Set
    End Property
    Public Property DocName() As String
        Get
            DocName = m_DocName
        End Get
        Set(ByVal value As String)
            m_DocName = value
        End Set
    End Property
    Public Property Document() As Object
        Get
            Document = m_Document
        End Get
        Set(ByVal value As Object)
            m_Document = value
        End Set
    End Property
    Public Property Document_Array() As Byte()
        Get
            Document_Array = m_Document_Array
        End Get
        Set(ByVal value As Byte())
            m_Document_Array = value
        End Set
    End Property
    Public Property Email() As Stream
        Get
            Email = m_Email
        End Get
        Set(ByVal value As Stream)
            m_Email = value
            Dim btFileGroupDescriptor As Byte() = New Byte(511) {}
            m_Email.Read(btFileGroupDescriptor, 0, 512)
            ' used to build the filename from the FileGroupDescriptor block

            Dim strFileName As New StringBuilder("")
            ' this trick gets the filename of the passed attached file

            Dim i As Integer = 76
            While btFileGroupDescriptor(i) <> 0
                strFileName.Append(Convert.ToChar(btFileGroupDescriptor(i)))
                i += 1
            End While

            m_DocName = strFileName.ToString
            m_DocDescription = Mid(strFileName.ToString, 1, m_DocName.Length - 4)
            m_DocFile = New FileInfo("c:\ExactTemp\" & strFileName.ToString)

            m_Email.Close()

            Dim theFile As String = m_DocFile.FullName

            'Dim ms As MemoryStream = DirectCast(e.Data.GetData("FileContents", True), MemoryStream)
            ' allocate enough bytes to hold the raw data

            Dim fileBytes As Byte() = New Byte(m_Contents.Length - 1) {}
            ' set starting position at first byte and read in the raw data

            m_Contents.Position = 0
            m_Contents.Read(fileBytes, 0, CInt(m_Contents.Length))
            ' create a file and save the raw zip file to it

            Dim fs As New FileStream(theFile, FileMode.Create)
            fs.Write(fileBytes, 0, CInt(fileBytes.Length))

            fs.Close()
            ' close the file

            Dim tempFile As New FileInfo(theFile)

            ' always good to make sure we actually created the file

            If tempFile.Exists = True Then
                'tempFile.Delete()
            Else
                'Debug.Print("File was not created!")
            End If

        End Set
    End Property
    Public ReadOnly Property GetDocumentText() As String
        Get
            GetDocumentText = String.Empty

            Dim dt As New DataTable
            Dim db As New cDBA
            Dim strSql As String

            strSql = _
             "SELECT   CAST(CAST(Document AS VARBINARY(MAX)) AS VARCHAR(MAX)) AS Text_Doc " & _
             "FROM	   Exact_Traveler_Document WITH (Nolock) " & _
             "WHERE    OnDrive = 2 AND DocID = " & m_DocID

            dt = db.DataTable(strSql) ' m_Sql(SqlMode.SelectNoLock))

            If dt.Rows.Count <> 0 Then
                GetDocumentText = dt.Rows(0).Item("Text_Doc").ToString.Trim
            End If

        End Get
    End Property
    Public ReadOnly Property GetDocumentImage() As Image
        Get
            Dim dt As New DataTable
            Dim db As New cDBA
            Dim strSql As String

            strSql = _
            "SELECT     * " & _
            "FROM       Exact_Traveler_Document WITH (Nolock) " & _
            "WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND Item_Guid = '" & m_Item_Guid & "' AND DocName = '" & SqlCompliantString(m_DocName) & "' AND DocTypeID = " & m_DocTypeID

            dt = db.DataTable(strSql) ' m_Sql(SqlMode.SelectNoLock))
            Dim oImageArray As Byte() = DirectCast(dt.Rows(0).Item("Document"), Byte())
            Dim oImage As Image = Nothing

            If Not oImageArray Is Nothing Then
                Using oMemStream As New MemoryStream(oImageArray, 0, oImageArray.Length)
                    oMemStream.Write(oImageArray, 0, oImageArray.Length)
                    oImage = Image.FromStream(oMemStream, True)
                End Using
            End If

            GetDocumentImage = oImage

        End Get
    End Property
    Public ReadOnly Property GetDocumentRow() As DataRow
        Get
            Dim dt As New DataTable
            Dim db As New cDBA

            If DocID <> 0 Then
                'Dim strSql As String = _
                '"SELECT * " & _
                '"FROM   Exact_Traveler_Document_Deleted WITH (Nolock) " & _
                '"WHERE  DocID = '" & m_DocID & "' "

                Dim strSql As String = _
                "SELECT * " & _
                "FROM   Exact_Traveler_Document WITH (Nolock) " & _
                "WHERE  DocID = '" & m_DocID & "' "
                dt = db.DataTable(strSql)
            Else
                '"FROM       OEI_Documents WITH (Nolock) " & _
                Dim strsql As String = _
                "SELECT     * " & _
                "FROM       Exact_Traveler_Document WITH (Nolock) " & _
                "WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND " & _
                "           DocName = '" & m_DocFile.Name & "' "

                '"SELECT     * " & _
                '"FROM       Exact_Traveler_Document WITH (Nolock) " & _
                '"WHERE      Ord_Guid = '" & m_Ord_Guid & "' AND " & _
                '"           Item_Guid = '" & m_Item_Guid & "' AND " & _
                '"           DocName = '" & m_DocFile.Name & "' "

                dt = db.DataTable(strsql)
            End If
            GetDocumentRow = dt.Rows(0)
        End Get
    End Property
    Public Property HeaderID() As Integer
        Get
            HeaderID = m_HeaderID
        End Get
        Set(ByVal value As Integer)
            m_HeaderID = value
        End Set
    End Property
    Public Property Item_Guid() As String
        Get
            Item_Guid = m_Item_Guid
        End Get
        Set(ByVal value As String)
            m_Item_Guid = value
            'Call CreateSql()
        End Set
    End Property
    Public Property OnDrive() As Integer
        Get
            OnDrive = m_OnDrive
        End Get
        Set(ByVal value As Integer)
            m_OnDrive = value
        End Set
    End Property
    Public Property Ord_Guid() As String
        Get
            Ord_Guid = m_Ord_Guid
        End Get
        Set(ByVal value As String)
            m_Ord_Guid = value
            'Call CreateSql()
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

End Class
