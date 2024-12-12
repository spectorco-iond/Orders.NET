Imports System.Xml
Imports System.IO

Public Class cOrderEDI

    Private m_OrdEDI_ID As Integer = 0
    Private m_EDI_Status As String '
    Private m_EDI_File_Name As String '
    Private m_XML_Text As String
    Private m_OEI_Status As String
    Private m_Ord_Guid As String
    Private m_Create_TS As Date
    'Private m_Note_Tm As Date
    'Private m_Document_Field_1 As Integer ' Can be busted, SQL says decimal 12,0
    'Private m_Filler_001 As String

#Region "Public constructors ##############################################"

    'Public Sub New()
    Public Sub New()

        Call Init()

    End Sub

    'Public Sub New(ByVal pID As Integer)
    Public Sub New(ByVal pID As Integer)

        Try

            Call Init()

            m_OrdEDI_ID = pID

            Call Load(pID)

        Catch er As Exception
            MsgBox("Error in cOrderEDI." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Sub New(ByVal pstrNote_Name As String, ByVal pstrNote_Name_Value As String)
    Public Sub New(ByVal pstrEDI_File_Name As String)

        Try

            Call Init()

            Call Load(pstrEDI_File_Name)
            'Call Load(pstrNote_Name, pstrNote_Name_Value)

        Catch er As Exception
            MsgBox("Error in cOrderEDI." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Private maintenance procedures ###################################"

    Private Sub Init()

        Try

            m_OrdEDI_ID = 0
            m_EDI_File_Name = String.Empty
            m_EDI_Status = String.Empty

            Call Reset()

        Catch er As Exception
            MsgBox("Error in cOrderEDI." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
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
            "FROM   OEI_ORDEDI O WITH (Nolock) " & _
            "WHERE  OrdEDI_ID = " & pintID & " "

            dt = db.DataTable(strSql)

            Call LoadLine(dt.Rows(0))

        Catch er As Exception
            MsgBox("Error in cOrderEDI." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '' Load is made private, it is loaded automatically on New. 
    Private Sub Load(ByVal pstrEDI_File_Name As String)

        Try
            Dim db As New cDBA
            Dim dt As New DataTable
            Dim strSql As String = String.Empty

            ' Exceptional. Signature cannot allow to have 2 identical so we force here
            ' IF IT IS A GUID, GET ID AND LOAD
            If pstrEDI_File_Name.ToUpper.Contains(".XML") Then

                ' ELSE ASSUMES IT IS A FILE NAME, GET ID AND LOAD.
                dt = db.DataTable("SELECT * FROM OEI_OrdEDI WITH (Nolock) WHERE EDI_File_Name = '" & pstrEDI_File_Name & "' ")
                'dt = db.DataTable("SELECT * FROM OEI_OrdEDI WITH (Nolock) WHERE EDI_File_Name = '" & pstrEDI_File_Name & "' AND ISNULL(EDI_Status, '') <> 'C' ")

                ' Load line if exists else create new data
                If dt.Rows.Count <> 0 Then
                    Load(CInt(dt.Rows(0).Item("OrdEDI_ID")))
                Else
                    m_EDI_File_Name = pstrEDI_File_Name
                    m_EDI_Status = "1" ' Started
                    m_OEI_Status = ""  ' Nothing yet
                    Save()
                End If

                dt = db.DataTable("SELECT * FROM OEI_Config WITH (Nolock) ")

                If dt.Rows.Count <> 0 Then

                    ' Create full file name
                    Dim strEDI_File As String = dt.Rows(0).Item("EDI_Files_Folder").ToString().Trim()
                    If strEDI_File(strEDI_File.Length - 1) <> "\" Then strEDI_File &= "\"
                    strEDI_File &= pstrEDI_File_Name

                    Dim strEDI_Backup As String = dt.Rows(0).Item("EDI_Backup_Folder").ToString().Trim()
                    If strEDI_Backup(strEDI_Backup.Length - 1) <> "\" Then strEDI_Backup &= "\"
                    strEDI_Backup &= pstrEDI_File_Name

                    ' Load document
                    Dim oDocument As XmlDocument = New XmlDocument()
                    oDocument.Load(strEDI_File)

                    m_EDI_Status = "E" ' Exported to SQL
                    m_XML_Text = oDocument.OuterXml

                    oDocument = Nothing

                    ' We only change status when empty. 
                    ' OEI can access the file even if trouble moving it to backyp folder.
                    If m_OEI_Status.Trim() = "" Then m_OEI_Status = "R" ' Ready to use in OEI
                    Save()

                    ' Backup the document
                    File.Copy(strEDI_File, strEDI_Backup, True)
                    File.Delete(strEDI_File)

                    m_EDI_Status = "C" ' EDI Import from service complete

                    'Dim oNode As XmlNode
                    'oNode = oDocument.SelectSingleNode("/ExactRequestSet/MacolaSalesOrderAddRequest/OrderHeader")

                    'Dim strOE_Po_No As String = String.Empty
                    'If Not oNode.SelectSingleNode("oe_po_no") Is Nothing Then strOE_Po_No = oNode.SelectSingleNode("oe_po_no").InnerText

                    'If strOE_Po_No <> String.Empty Then
                    '    Dim dtCheckException As DataTable
                    '    strSql = "SELECT * FROM OEI_ORD_EXCEPTION WITH (Nolock) WHERE OE_PO_No = '" & strOE_Po_No.Replace("'", "''") & "' "
                    '    dtCheckException = db.DataTable(strSql)

                    '    If dtCheckException.Rows.Count <> 0 Then

                    '        ' When unblocking an order
                    '        ' set m_EDI_Status to C and m_OEI_Status to R

                    '        m_EDI_Status = "B" ' EDI order completed and BLOCKED
                    '        m_OEI_Status = "B" ' Exported to SQL

                    '        strSubject = "***** EDI ORDER BLOCKED BY HARVEST IMPORT CONTROL *****"
                    '        strMessage = _
                    '        "Order with PO Number " & strOE_Po_No & " was blocked by Harvest Import control.<br>" & _
                    '        "The order number was found in table OEI_EDI_EXCEPTION."
                    '        Call SendEmail()
                    '    End If
                    'End If

                    Save()

                End If

            Else

                dt = db.DataTable("SELECT * FROM OEI_OrdEDI WITH (Nolock) WHERE Ord_Guid = '" & pstrEDI_File_Name & "' ")

                If dt.Rows.Count <> 0 Then
                    Load(CInt(dt.Rows(0).Item("OrdEDI_ID")))
                End If

            End If

        Catch er As Exception
            MsgBox("Error in cOrderEDI." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Loads the note fields from the DataRow record
    Private Sub LoadLine(ByRef pdrRow As DataRow)

        Try
            ' Loads the item from the DataRow into the locals
            If Not (pdrRow.Item("OrdEDI_ID").Equals(DBNull.Value)) Then m_OrdEDI_ID = pdrRow.Item("OrdEDI_ID")
            If Not (pdrRow.Item("EDI_File_Name").Equals(DBNull.Value)) Then m_EDI_File_Name = pdrRow.Item("EDI_File_Name").ToString
            If Not (pdrRow.Item("EDI_Status").Equals(DBNull.Value)) Then m_EDI_Status = pdrRow.Item("EDI_Status").ToString
            If Not (pdrRow.Item("XML_Text").Equals(DBNull.Value)) Then m_XML_Text = pdrRow.Item("XML_Text").ToString
            If Not (pdrRow.Item("OEI_Status").Equals(DBNull.Value)) Then m_OEI_Status = pdrRow.Item("OEI_Status").ToString
            If Not (pdrRow.Item("Ord_Guid").Equals(DBNull.Value)) Then m_Ord_Guid = pdrRow.Item("Ord_Guid").ToString
            If Not (pdrRow.Item("Create_TS").Equals(DBNull.Value)) Then m_Create_TS = pdrRow.Item("Create_TS")

        Catch er As Exception
            MsgBox("Error in cOrderEDI." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Saves the note fields into the DataRow record
    Private Sub SaveLine(ByRef pdrRow As DataRow)

        Try

            ' Save the locals into the datarow.
            pdrRow.Item("EDI_File_Name") = m_EDI_File_Name
            pdrRow.Item("EDI_Status") = m_EDI_Status
            pdrRow.Item("XML_Text") = m_XML_Text
            pdrRow.Item("OEI_Status") = m_OEI_Status
            pdrRow.Item("Ord_Guid") = m_Ord_Guid
            If m_OrdEDI_ID = 0 Then
                pdrRow.Item("Create_TS") = Date.Now.Date
            End If

            ' Don't resave ID, it saves as is or create a new one.
            'pdrRow.Item("OrdEDI_ID") = m_OrdEDI_ID

        Catch er As Exception
            MsgBox("Error in cOrderEDI." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public maintenance procedures ####################################"

    ' Deletes the current note from the database, if it exists.
    Public Sub Delete()

        Try

            If m_OrdEDI_ID = 0 Then Exit Sub

            Dim db As New cDBA
            Dim dt As New DataTable

            Dim strSql As String

            strSql = _
            "DELETE FROM OEI_OrdEDI " & _
            "WHERE  OrdEDI_ID = " & m_OrdEDI_ID & " "

            db.Execute(strSql)

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in cOrderEDI." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Reset every non mandatory field of a note
    Public Sub Reset()

        ' Resets the locals to empty values, all but ID, Note_Name and Note_Name_Value
        m_XML_Text = String.Empty
        m_OEI_Status = String.Empty ' g_User.Lock_Code ' String.Empty
        m_Ord_Guid = String.Empty
        m_Create_TS = New Date

    End Sub

    ' Update the current note into the database, or creates it if not existing
    Public Sub Save()

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Try

            Dim db As New cDBA
            Dim dt As New DataTable
            Dim drRow As DataRow

            Dim strSql As String

            strSql = _
                    "SELECT * " & _
                    "FROM   OEI_OrdEDI O " & _
                    "WHERE  OrdEDI_ID = " & m_OrdEDI_ID & " "

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

            If m_OrdEDI_ID = 0 Then

                strSql = _
                        "SELECT * " & _
                        "FROM   OEI_OrdEDI O " & _
                        "WHERE  EDI_File_Name = '" & m_EDI_File_Name & "' "
                dt = db.DataTable(strSql)

                If dt.Rows.Count <> 0 Then
                    m_OrdEDI_ID = dt.Rows(0).Item("ORDEDI_ID")
                End If

            End If

        Catch er As Exception
            MsgBox("Error in cOrderEDI." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region

#Region "Public properties ################################################"

    Public Property OrdEDI_ID() As Integer
        Get
            OrdEDI_ID = m_OrdEDI_ID
        End Get
        Set(ByVal value As Integer)
            m_OrdEDI_ID = value
        End Set
    End Property
    Public Property EDI_File_Name() As String
        Get
            EDI_File_Name = m_EDI_File_Name
        End Get
        Set(ByVal value As String)
            m_EDI_File_Name = value
        End Set
    End Property
    Public Property EDI_Status() As String
        Get
            EDI_Status = m_EDI_Status
        End Get
        Set(ByVal value As String)
            m_EDI_Status = value
        End Set
    End Property
    Public Property XML_Text() As String
        Get
            XML_Text = m_XML_Text
        End Get
        Set(ByVal value As String)
            m_XML_Text = value
        End Set
    End Property
    Public Property OEI_Status() As String
        Get
            OEI_Status = m_OEI_Status
        End Get
        Set(ByVal value As String)
            m_OEI_Status = value
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
    Public Property Create_TS() As Date
        Get
            Create_TS = m_Create_TS
        End Get
        Set(ByVal value As Date)
            m_Create_TS = value
        End Set
    End Property

#End Region

End Class
