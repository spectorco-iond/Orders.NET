Imports System.Collections.Generic

Public Class cContact

    ' Information from Cicntp
    Private m_ID As Integer = 0
    Private m_Cnt_ID As String = ""
    'Private m_Cmp_WWN As String = ""
    Private m_Gender As String = ""
    Private m_PredCode As String = ""
    Private m_FirstName As String = ""
    Private m_LastName As String = ""
    Private m_FullName As String = ""
    Private m_TaalCode As String = ""
    Private m_Cnt_Acc_Man As String = ""
    Private m_Cnt_F_Tel As String = ""
    Private m_Cnt_F_Ext As String = ""
    Private m_Cnt_F_Fax As String = ""
    Private m_Cnt_Email As String = ""

    Private m_OrderContact As COrderContact
    Private m_Customer As cCustomer
    Public Shared popMessageLog As New Dictionary(Of String, String)

    '' Information for traveler order
    'Public Cus_No As String = ""
    'Public OrderNo As String = ""
    'Public DefContact As Integer = 0
    'Private m_DefMethod As String = "0"
    'Public ContactType As Integer = 0
    'Private m_PrimaryContact As Integer = 0
    'Public DateAdded As Date
    'Public Ord_GUID As String = ""

    Public Class COrderContact

        ' Information for traveler order
        Private m_Cus_No As String = ""
        Private m_OrderNo As String = ""
        Private m_ID As Integer = 0
        Private m_DefMethod As ContactMethod = ContactMethod.None ' String = "0"
        Private m_ContactType As Integer = 0
        Private m_PrimaryContact As Integer = 0
        Private m_DateAdded As Date
        Private m_Ord_GUID As String = ""

        Private m_Contact As New cContact

        Public Sub New()

        End Sub

        'Public Sub New(ByVal pstrOrd_No As String, ByVal pintContactType As ContactType)

        '    Try

        '        If IsNumeric(pstrOrd_No) Then
        '            m_OrderNo = pstrOrd_No
        '        Else
        '            m_Ord_GUID = pstrOrd_No
        '        End If

        '        ContactType = pintContactType

        '        Call Load()

        '    Catch er As Exception
        '        MsgBox("Error in CContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        '    End Try

        'End Sub

        Public Sub New(ByVal pstrOrd_No As String, ByVal pintContactType As ContactType, ByVal pContact As cContact)

            Try

                m_Contact = pContact

                If IsNumeric(pstrOrd_No) Then
                    m_OrderNo = pstrOrd_No
                Else
                    m_Ord_GUID = pstrOrd_No
                End If

                ContactType = pintContactType

                Call Load()

            Catch er As Exception
                MsgBox("Error in CContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        Public Sub New(ByVal pintContactID As Integer)

            Try

                m_ID = pintContactID

            Catch er As Exception
                MsgBox("Error in CContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        Public Sub Delete()

            Dim conn As New cDBA
            Dim dt As New DataTable
            Dim strSql As String

            'If IsNumeric(OrderNo) Then
            '    strSql = "" & _
            '    "SELECT     * " & _
            '    "FROM       EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW " & _
            '    "WHERE      RTRIM(LTRIM(OrderNo)) = '" & Trim(OrderNo) & "' AND " & _
            '    "           DefContact = " & ID & " AND " & _
            '    "           ContactType = " & ContactType
            'Else
            '    strSql = "" & _
            '    "SELECT     * " & _
            '    "FROM       EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW " & _
            '    "WHERE      RTRIM(LTRIM(Ord_Guid)) = '" & Trim(Ord_Guid) & "' AND " & _
            '    "           DefContact = " & ID & " AND " & _
            '    "           ContactType = " & ContactType
            'End If

            Try
                If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

                strSql = "" &
                "SELECT     * " &
                "FROM       OEI_ORDER_CONTACTS " &
                "WHERE      RTRIM(LTRIM(Ord_Guid)) = '" & Trim(Ord_GUID) & "' AND " &
                "           DefContact = " & ID & " AND " &
                "           ContactType = " & ContactType

                dt = conn.DataTable(strSql)
                If dt.Rows.Count <> 0 Then
                    conn.DBDataTable.Rows(0).Delete()
                    conn.DBDataAdapter.Update(conn.DBDataTable)
                End If

            Catch oe_er As OEException
                MsgBox(oe_er.Message)
            Catch er As Exception
                MsgBox("Error in COrderContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        Public Sub Load()

            Try

                'If ID = 0 Then Exit Sub

                Dim strSql As String
                Dim conn As New cDBA
                Dim dt As DataTable

                'strSql = _
                '"SELECT     * " & _
                '"FROM       EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW with (nolock) " & _
                '"WHERE      OrderNo = '" & OrderNo & "' AND ContactType = " & ContactType

                strSql =
                "SELECT     * " &
                "FROM       OEI_ORDER_CONTACTS with (nolock) " &
                "WHERE      Ord_Guid = '" & Ord_GUID & "' AND ContactType = " & ContactType

                dt = conn.DataTable(strSql)
                If dt.Rows.Count <> 0 Then

                    'Cus_No = IIf(dt.Rows(0).Item("Cus_No").Equals(DBNull.Value), "", dt.Rows(0).Item("Cus_No"))
                    ID = IIf(dt.Rows(0).Item("DefContact").Equals(DBNull.Value), "", dt.Rows(0).Item("DefContact"))
                    m_DefMethod = ContactMethodInt(IIf(dt.Rows(0).Item("DefMethod").Equals(DBNull.Value), "0", dt.Rows(0).Item("DefMethod")))
                    m_PrimaryContact = 1 ' IIf(dt.Rows(0).Item("PrimaryContact").Equals(DBNull.Value), 0, dt.Rows(0).Item("PrimaryContact"))
                    'DateAdded = IIf(dt.Rows(0).Item("DateAdded").Equals(DBNull.Value), "", dt.Rows(0).Item("DateAdded"))

                End If
                dt.Dispose()

            Catch er As Exception
                MsgBox("Error in CContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        Public Sub Reset()

            m_Cus_No = String.Empty
            m_OrderNo = String.Empty
            m_ID = 0
            m_DefMethod = ContactMethod.None ' String = "0"
            m_ContactType = 0
            m_PrimaryContact = 0
            m_DateAdded = Now()
            m_Ord_GUID = String.Empty

            m_Contact = New cContact

        End Sub

        Public Sub Save()

            If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

            Try
                Dim conn As New cDBA
                Dim dt As New DataTable
                Dim drDataRow As DataRow

                'Dim strSql As String = _
                '"SELECT     * " & _
                '"FROM       OEI_ORDER_CONTACTS " & _
                '"WHERE      RTRIM(LTRIM(Ord_Guid)) = '" & Trim(Ord_GUID) & "' AND " & _
                '"           DefContact = " & ID & " AND " & _
                '"           ContactType = " & ContactType

                Dim strSql As String =
                "SELECT     * " &
                "FROM       OEI_ORDER_CONTACTS " &
                "WHERE      RTRIM(LTRIM(Ord_Guid)) = '" & Trim(Ord_GUID) & "' AND " &
                "           ContactType = " & ContactType

                dt = conn.DataTable(strSql)
                If dt.Rows.Count <> 0 Then
                    drDataRow = dt.Rows(0)
                Else
                    drDataRow = dt.NewRow
                End If

                'drDataRow.Item("Cus_No") = Cus_No
                'drDataRow.Item("OrderNo") = OrderNo
                drDataRow.Item("Ord_Guid") = Ord_GUID
                If m_ID <> 0 Then
                    drDataRow.Item("DefContact") = m_ID ' m_Contact.ID
                Else
                    drDataRow.Item("DefContact") = m_Contact.ID ' m_Contact.ID
                End If
                'If m_DefMethod = ContactMethod.None Then
                Dim strSendMethod As String

                strSendMethod = Trim(LastSendMethod())
                If strSendMethod = "" Then
                    If m_Contact.Cnt_Email <> "" Then
                        m_DefMethod = ContactMethod.Email
                    ElseIf m_Contact.Cnt_F_Fax <> "" Then
                        m_DefMethod = ContactMethod.Fax
                    Else
                        m_DefMethod = ContactMethod.Email
                    End If
                Else
                    m_DefMethod = ContactMethodInt(strSendMethod)
                End If
                'End If

                drDataRow.Item("DefMethod") = ContactMethodString(m_DefMethod)
                drDataRow.Item("ContactType") = ContactType
                'drDataRow.Item("PrimaryContact") = m_PrimaryContact
                'drDataRow.Item("DateAdded") = DateAdded

                'Debug.Print("Ord_Guid: char   " & drDataRow.Item("Ord_Guid"))
                'Debug.Print("DefContact: int  " & drDataRow.Item("DefContact"))
                'Debug.Print("DefMethod: char  " & drDataRow.Item("DefMethod"))
                'Debug.Print("ContactType: int " & drDataRow.Item("ContactType"))

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
                MsgBox("Error in COrderContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Sub

        Public Function LastSendMethod() As String

            LastSendMethod = "E"

            'If m_ID = 0 Then Exit Function

            Try
                ' 402595  403669 
                'Dim strSql As String = "" & _
                '"SELECT ISNULL(DefMethod, '') AS DefMethod " & _
                '"FROM EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW WITH (Nolock) " & _
                '"WHERE DefContact = " & ID & " AND RTRIM(LTRIM(OrderNo)) <> '" & Trim(OrderNo) & "' " & _
                '"ORDER BY DateAdded Desc "

                'Dim strSql As String = "" & _
                '"SELECT ISNULL(DefMethod, '') AS DefMethod " & _
                '"FROM OEI_ORDER_CONTACTS WITH (Nolock) " & _
                '"WHERE DefContact = " & ID & " AND RTRIM(LTRIM(Ord_Guid)) <> '" & Trim(Ord_GUID) & "' " & _
                '"ORDER BY DateAdded Desc "

                Dim strSql As String = "" &
                "SELECT     ISNULL(DefMethod, '') AS DefMethod, DateAdded, OrderNo, 1 AS ID " &
                "FROM       EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW WITH (Nolock) " &
                "WHERE      DefContact = " & ID & " " &
                "UNION " &
                "SELECT     ISNULL(DefMethod, '') AS DefMethod, H.Ord_Dt AS DateAdded, H.Ord_No AS OrderNo, C.ID " &
                "FROM       OEI_ORDER_CONTACTS C WITH (Nolock) " &
                "INNER JOIN OEI_ORDHDR H WITH (Nolock) ON C.Ord_Guid = H.Ord_Guid " &
                "WHERE      DefContact = " & ID & " AND RTRIM(LTRIM(C.Ord_Guid)) <> '" & Trim(Ord_GUID) & "' " &
                "ORDER BY   DateAdded Desc, ID DESC "

                Dim conn As New cDBA
                Dim dt As New DataTable
                dt = conn.DataTable(strSql)
                If dt.Rows.Count <> 0 Then
                    LastSendMethod = dt.Rows(0).Item("DefMethod")
                    If LastSendMethod = "F" Then
                        m_DefMethod = ContactMethod.Fax
                    Else
                        m_DefMethod = ContactMethod.Email
                    End If
                    LastSendMethod = IIf(m_DefMethod = ContactMethod.Fax, "F", "E") '   DefMethodString(m_DefMethod)
                    'LastSendMethod = "E"
                End If

                dt.Dispose()

            Catch er As Exception
                MsgBox("Error in COrderContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
            End Try

        End Function

#Region "Public properties ############################################"

        Public Property ContactType() As Integer
            Get
                ContactType = m_ContactType
            End Get
            Set(ByVal value As Integer)
                m_ContactType = value
            End Set
        End Property
        Public Property Contact() As cContact
            Get
                Contact = m_Contact
            End Get
            Set(ByVal value As cContact)
                m_Contact = value
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
        Public Property DateAdded() As Date
            Get
                DateAdded = m_DateAdded
            End Get
            Set(ByVal value As Date)
                m_DateAdded = value
            End Set
        End Property
        Public Property DefMethod() As ContactMethod
            Get
                DefMethod = m_DefMethod
            End Get
            Set(ByVal value As ContactMethod)
                m_DefMethod = value
            End Set
        End Property
        Public ReadOnly Property DefMethodString() As String
            Get
                DefMethodString = ""
                Select Case m_DefMethod
                    Case ContactMethod.None
                        DefMethodString = ""
                    Case ContactMethod.Email
                        DefMethodString = "E"
                    Case ContactMethod.Fax
                        DefMethodString = "F"
                End Select
                DefMethodString = ""
            End Get
        End Property
        'Public Function DefMethodString(ByVal p) As String
        '    Get
        '    DefMethodString = ""
        '    Select Case m_DefMethod
        '        Case ContactMethod.None
        '            DefMethodString = ""
        '        Case ContactMethod.Email
        '            DefMethodString = "E"
        '        Case ContactMethod.Fax
        '            DefMethodString = "F"
        '    End Select
        '    DefMethodString = ""
        'End Function
        ''Set(ByVal value As String)
        '    If m_ID = 0 Then
        '        m_DefMethod = ContactMethod.None
        '        Exit Property
        '    End If
        '    Select Case value
        '        Case "E"
        '            m_DefMethod = ContactMethod.Email
        '        Case "F"
        '            m_DefMethod = ContactMethod.Fax
        '        Case Else
        '            m_DefMethod = ContactMethod.None
        '    End Select
        'End Set
        'End Property
        Public Property ID() As Integer
            Get
                ID = m_ID
            End Get
            Set(ByVal value As Integer)
                m_ID = value
            End Set
        End Property
        Public Property Ord_No() As String
            Get
                Ord_No = m_OrderNo
            End Get
            Set(ByVal value As String)
                m_OrderNo = value
            End Set
        End Property
        Public Property Ord_GUID() As String
            Get
                Ord_GUID = m_Ord_GUID
            End Get
            Set(ByVal value As String)
                m_Ord_GUID = value
            End Set
        End Property
        Public Property PrimaryContact() As Integer
            Get
                PrimaryContact = m_PrimaryContact
            End Get
            Set(ByVal value As Integer)
                If ID = 0 Then
                    m_PrimaryContact = 0
                    Exit Property
                End If
                m_PrimaryContact = value
            End Set
        End Property

#End Region

    End Class

    ' Shall never be used !!
    Public Sub New()

    End Sub

    Public Sub New(ByVal pstrOrd_No As String, ByVal pintContactType As ContactType)

        Try

            m_OrderContact = New COrderContact(pstrOrd_No, pintContactType, Me)
            m_ID = m_OrderContact.ID

            Call Load()

        Catch er As Exception
            MsgBox("Error in CContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Sub New(ByVal pintContactID As Integer)

    '    Try

    '        m_ID = pintContactID

    '        Call Load()

    '        m_OrderContact = New COrderContact()

    '        m_OrderContact.DefMethod = m_OrderContact.LastSendMethod

    '    Catch er As Exception
    '        MsgBox("Error in CContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    '    'm_ID = pintContactID
    '    'Call Load()

    'End Sub

    Public Sub Delete()

        'Dim conn As New cDBA
        'Dim dt As New DataTable
        'Dim strSql As String

        ''If IsNumeric(OrderNo) Then
        ''    strSql = "" & _
        ''    "SELECT     * " & _
        ''    "FROM       EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW " & _
        ''    "WHERE      RTRIM(LTRIM(OrderNo)) = '" & Trim(OrderNo) & "' AND " & _
        ''    "           DefContact = " & ID & " AND " & _
        ''    "           ContactType = " & ContactType
        ''Else
        ''    strSql = "" & _
        ''    "SELECT     * " & _
        ''    "FROM       EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW " & _
        ''    "WHERE      RTRIM(LTRIM(Ord_Guid)) = '" & Trim(Ord_Guid) & "' AND " & _
        ''    "           DefContact = " & ID & " AND " & _
        ''    "           ContactType = " & ContactType
        ''End If

        'strSql = "" & _
        '"SELECT     * " & _
        '"FROM       OEI_ORDER_CONTACTS " & _
        '"WHERE      RTRIM(LTRIM(Ord_Guid)) = '" & Trim(Ord_GUID) & "' AND " & _
        '"           DefContact = " & ID & " AND " & _
        '"           ContactType = " & ContactType

        'dt = conn.DataTable(strSql)
        'If dt.Rows.Count <> 0 Then
        '    conn.DBDataTable.Rows(0).Delete()
        '    conn.DBDataAdapter.Update(conn.DBDataTable)
        'End If

    End Sub

    Public Sub Init()

        ' Information from Cicntp
        m_ID = 0
        m_Cnt_ID = String.Empty

        m_Gender = String.Empty
        m_PredCode = String.Empty
        m_FirstName = String.Empty
        m_LastName = String.Empty
        m_FullName = String.Empty
        m_TaalCode = String.Empty
        m_Cnt_Acc_Man = String.Empty
        m_Cnt_F_Tel = String.Empty
        m_Cnt_F_Ext = String.Empty
        m_Cnt_F_Fax = String.Empty
        m_Cnt_Email = String.Empty

        m_OrderContact = New COrderContact
        m_Customer = New cCustomer

    End Sub

    Public Sub Load()

        Try

            If m_ID = 0 Then Exit Sub

            Dim strSql As String
            Dim conn As New cDBA
            Dim dt As DataTable

            strSql = " SELECT * FROM CiCntp WITH (nolock) where ISNULL(active_y,0) = 1 And ID = " & ID

            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 
            dt = conn.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                m_Cnt_ID = dt.Rows(0).Item("Cnt_ID").ToString
                m_Gender = IIf(dt.Rows(0).Item("Gender").Equals(DBNull.Value), "", dt.Rows(0).Item("Gender"))
                m_PredCode = IIf(dt.Rows(0).Item("PredCode").Equals(DBNull.Value), "", dt.Rows(0).Item("PredCode"))
                m_TaalCode = IIf(dt.Rows(0).Item("TaalCode").Equals(DBNull.Value), "", dt.Rows(0).Item("TaalCode"))
                m_Cnt_F_Tel = IIf(dt.Rows(0).Item("Cnt_F_Tel").Equals(DBNull.Value), "", dt.Rows(0).Item("Cnt_F_Tel"))
                m_Cnt_F_Ext = IIf(dt.Rows(0).Item("Cnt_F_Ext").Equals(DBNull.Value), "", dt.Rows(0).Item("Cnt_F_Ext"))
                m_Cnt_F_Fax = IIf(dt.Rows(0).Item("Cnt_F_Fax").Equals(DBNull.Value), "", dt.Rows(0).Item("Cnt_F_Fax"))
                m_Cnt_Email = IIf(dt.Rows(0).Item("Cnt_Email").Equals(DBNull.Value), "", dt.Rows(0).Item("Cnt_Email"))
                m_Cnt_Acc_Man = IIf(dt.Rows(0).Item("Cnt_Acc_Man").Equals(DBNull.Value), "", dt.Rows(0).Item("Cnt_Acc_Man"))
                m_FirstName = IIf(dt.Rows(0).Item("Cnt_F_Name").Equals(DBNull.Value), "", dt.Rows(0).Item("Cnt_F_Name"))
                m_LastName = IIf(dt.Rows(0).Item("Cnt_L_Name").Equals(DBNull.Value), "", dt.Rows(0).Item("Cnt_L_Name"))
                m_FullName = IIf(dt.Rows(0).Item("FullName").Equals(DBNull.Value), "", dt.Rows(0).Item("FullName"))

            End If
            dt.Dispose()

        Catch er As Exception
            MsgBox("Error in CContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Reset()

        m_Gender = String.Empty
        m_PredCode = String.Empty
        m_FirstName = String.Empty
        m_LastName = String.Empty
        m_FullName = String.Empty
        m_TaalCode = String.Empty
        m_Cnt_Acc_Man = String.Empty
        m_Cnt_F_Tel = String.Empty
        m_Cnt_F_Ext = String.Empty
        m_Cnt_F_Fax = String.Empty
        m_Cnt_Email = String.Empty

        m_OrderContact = New COrderContact
        'm_Customer = New cOrdHead.cCustomer

    End Sub

    Public Sub Save()

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Dim strSql As String
        Dim dt As DataTable
        Dim db As New cDBA

        Try
            strSql = " INSERT INTO cicntp ( cmp_wwn, cnt_f_name, cnt_l_name, cnt_m_name, FullName, Gender, predcode, cnt_job_desc, " & " taalcode, cnt_f_ext, cnt_f_fax, cnt_f_tel, cnt_f_mobile, cnt_email, cnt_acc_man, Administration, syscreator, " & " sysmodifier, numberfield1, numberfield2, numberfield3, numberfield4, numberfield5 ) " & " VALUES ( '" & m_Customer.cmp_wwn & "', '" & SqlCompliantString(m_FirstName) & "', '" & SqlCompliantString(m_LastName) & "', '', " & " '" & SqlCompliantString(m_FullName) & "', '" & m_Gender & "', '" & m_PredCode & "', " & " '', '" & m_TaalCode & "', '" & SqlCompliantString(m_Cnt_F_Ext) & "', '" & SqlCompliantString(m_Cnt_F_Fax) & "', " & " '" & SqlCompliantString(m_Cnt_F_Tel) & "', '', '" & SqlCompliantString(m_Cnt_Email) & "', '" & SqlCompliantString(m_Cnt_Acc_Man) & "', " & " '100', '1', '1', '0', '0', '0', '0', '0' ) "
            db.Execute(strSql)

            'If gConn.State <> ADODB.ObjectStateEnum.adStateOpen Then gConn.Open()
            'gConn.Execute(strSql)
            'If gConn.State <> ADODB.ObjectStateEnum.adStateClosed Then gConn.Close()

            ' Retrieve Contact info
            strSql = " SELECT * FROM cicntp WITH (Nolock) WHERE ISNULL(active_y,0) = 1 And cmp_wwn = '" & m_Customer.cmp_wwn & "' ORDER BY ID DESC "
            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 
            dt = db.DataTable(strSql)
            'Dim rsContactInfo As New ADODB.Recordset
            'rsContactInfo.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)


            'If gConn.State <> ADODB.ObjectStateEnum.adStateOpen Then gConn.Open()

            ' Add Contact info to Addresses, ShippingDays and BacoDiscussionStdPictures tables
            'strSql = " INSERT INTO Addresses ( Type, Account, ContactPerson, AddressLine1, AddressLine2, AddressLine3, " & " PostCode, City, StateCode, County, Country, KeepSameAsVisit, textfield1, Main ) " & " VALUES ( 'DEL', '" & m_Customer.Cmp_WWN & "', '" & rsContactInfo.Fields("cnt_id").Value & "', '" & SqlCompliantString(m_Customer.Cmp_FAdd1) & "', " & " '" & SqlCompliantString(m_Customer.Cmp_FAdd2) & "', '" & SqlCompliantString(m_Customer.Cmp_FAdd3) & "', '" & SqlCompliantString(m_Customer.Cmp_FPC) & "', " & " '" & SqlCompliantString(m_Customer.Cmp_FCity) & "', '" & SqlCompliantString(m_Customer.Cmp_FCounty) & "', '" & m_Customer.StateCode & "', " & " '" & m_Customer.Cmp_FCtry & "', '1', '" & SqlCompliantString(m_Customer.Cmp_Name) & "', '0' ) "
            strSql = " INSERT INTO Addresses ( Type, Account, ContactPerson, AddressLine1, AddressLine2, AddressLine3, " & " PostCode, City, StateCode, County, Country, KeepSameAsVisit, textfield1, Main ) " & " VALUES ( 'DEL', '" & m_Customer.cmp_wwn & "', '" & dt.Rows(0).Item("cnt_id").ToString & "', '" & SqlCompliantString(m_Customer.cmp_fadd1) & "', " & " '" & SqlCompliantString(m_Customer.cmp_fadd2) & "', '" & SqlCompliantString(m_Customer.cmp_fadd3) & "', '" & SqlCompliantString(m_Customer.cmp_fpc) & "', " & " '" & SqlCompliantString(m_Customer.cmp_fcity) & "', '" & SqlCompliantString(m_Customer.cmp_fcounty) & "', '" & m_Customer.StateCode & "', " & " '" & m_Customer.cmp_fctry & "', '1', '" & SqlCompliantString(m_Customer.cmp_name) & "', '0' ) "
            db.Execute(strSql)
            'gConn.Execute(strSql)

            'strSql = " INSERT INTO Addresses ( Type, Account, ContactPerson, AddressLine1, AddressLine2, AddressLine3, " & " PostCode, City, StateCode, County, Country, KeepSameAsVisit, textfield1, Main ) " & " VALUES ( 'INV', '" & m_Customer.Cmp_WWN & "', '" & rsContactInfo.Fields("cnt_id").Value & "', '" & sqlcompliantstring(m_Customer.Cmp_FAdd1) & "', " & " '" & sqlcompliantstring(m_Customer.Cmp_FAdd2) & "', '" & sqlcompliantstring(m_Customer.Cmp_FAdd3) & "', '" & sqlcompliantstring(m_Customer.Cmp_FPC) & "', " & " '" & sqlcompliantstring(m_Customer.Cmp_FCity) & "', '" & sqlcompliantstring(m_Customer.Cmp_FCounty) & "', '" & m_Customer.StateCode & "', " & " '" & m_Customer.Cmp_FCtry & "', '1', '" & sqlcompliantstring(m_Customer.Cmp_Name) & "', '0' ) "
            'gConn.Execute(strSql)

            ' strSql = " INSERT INTO Addresses ( Type, Account, ContactPerson, AddressLine1, AddressLine2, AddressLine3, " & " PostCode, City, StateCode, County, Country, KeepSameAsVisit, textfield1, Main ) " & " VALUES ( 'POS', '" & m_Customer.cmp_wwn & "', '" & dt.Rows(0).Item("cnt_id").ToString & "', '" & SqlCompliantString(m_Customer.cmp_fadd1) & "', " & " '" & SqlCompliantString(m_Customer.cmp_fadd2) & "', '" & SqlCompliantString(m_Customer.cmp_fadd3) & "', '" & SqlCompliantString(m_Customer.cmp_fpc) & "', " & " '" & SqlCompliantString(m_Customer.cmp_fcity) & "', '" & SqlCompliantString(m_Customer.cmp_fcounty) & "', '" & m_Customer.StateCode & "', " & " '" & m_Customer.cmp_fctry & "', '1', '" & SqlCompliantString(m_Customer.cmp_name) & "', '0' ) "
            'db.Execute(strSql)
            'gConn.Execute(strSql)

            'strSql = " INSERT INTO Addresses ( Type, Account, ContactPerson, AddressLine1, AddressLine2, AddressLine3, " & " PostCode, City, StateCode, County, Country, KeepSameAsVisit, textfield1, Main ) " & " VALUES ( 'VIS', '" & m_Customer.Cmp_WWN & "', '" & rsContactInfo.Fields("cnt_id").Value & "', '" & sqlcompliantstring(m_Customer.Cmp_FAdd1) & "', " & " '" & sqlcompliantstring(m_Customer.Cmp_FAdd2) & "', '" & sqlcompliantstring(m_Customer.Cmp_FAdd3) & "', '" & sqlcompliantstring(m_Customer.Cmp_FPC) & "', " & " '" & sqlcompliantstring(m_Customer.Cmp_FCity) & "', '" & sqlcompliantstring(m_Customer.Cmp_FCounty) & "', '" & m_Customer.StateCode & "', " & " '" & m_Customer.Cmp_FCtry & "', '1', '" & sqlcompliantstring(m_Customer.Cmp_Name) & "', '0' ) "
            'gConn.Execute(strSql)

            'If gConn.State <> ADODB.ObjectStateEnum.adStateClosed Then gConn.Close()

            strSql = " SELECT ID " & " FROM Addresses " & " WHERE Account = '" & m_Customer.cmp_wwn & "'" & " AND ContactPerson = '" & dt.Rows(0).Item("cnt_id").ToString & "' " & " ORDER BY ID DESC "

            'Dim rsAddressInfo As New ADODB.Recordset
            Dim dtAddressInfo As DataTable
            dtAddressInfo = db.DataTable(strSql)
            'rsAddressInfo.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            'If gConn.State <> ADODB.ObjectStateEnum.adStateOpen Then gConn.Open()

            strSql = " INSERT INTO ShippingDays ( AddressID ) " & " VALUES ( '" & dtAddressInfo.Rows(0).Item("ID").ToString & "' ) "
            db.Execute(strSql)
            'gConn.Execute(strSql)

            strSql = " INSERT INTO BacoDiscussionStdPictures ( Description, Filename, Type, Width, Height, Size, Division ) " & " VALUES ( '_" & SqlCompliantString(m_Customer.cmp_code) & "', '" & SqlCompliantString(m_FullName) & "', 'vCard', " & " '1', '1', '1', '100' ) "
            db.Execute(strSql)
            'gConn.Execute(strSql)



            strSql = " INSERT INTO EXACT_TRAVELER_NEW_CONTACTS_FROM_APP ( cnt_id, cmp_wwn, AddingUser ) " & " VALUES ( '" & dt.Rows(0).Item("cnt_id").ToString & "', '" & m_Customer.cmp_wwn & "', '" & SqlCompliantString(GetNTUserID()) & "' ) "
            db.Execute(strSql)
            'gConn.Execute(strSql)

            'rsAddressInfo.Close()
            'rsContactInfo.Close()

            'If gConn.State <> ADODB.ObjectStateEnum.adStateClosed Then gConn.Close()

        Catch er As Exception
            MsgBox("Error in CContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Save_ESynergy(ByVal cnt_id As String)
        'Call Save_ESynergy(dt.Rows(0).Item("cnt_id").ToString)
        Dim strSql As String = ""
        Dim dtAddressInfo As DataTable

        Try
            'connection to exact_ese db
            Dim db As New cDBA(g_Default_ConnSynergy_String)
            strSql = "INSERT INTO cicntp (cnt_id, cmp_wwn, cnt_f_name, cnt_l_name, cnt_m_name, FullName, Gender, predcode, cnt_job_desc, " & " taalcode, cnt_f_ext, cnt_f_fax, cnt_f_tel, cnt_f_mobile, cnt_email, cnt_acc_man, Administration, syscreator, " & " sysmodifier, numberfield1, numberfield2, numberfield3, numberfield4, numberfield5 ) " & " VALUES ('" & m_Customer.cnt_id & "', '" & m_Customer.cmp_wwn & "', '" & SqlCompliantString(m_FirstName) & "', '" & SqlCompliantString(m_LastName) & "', '', " & " '" & SqlCompliantString(m_FullName) & "', '" & m_Gender & "', '" & m_PredCode & "', " & " '', '" & m_TaalCode & "', '" & SqlCompliantString(m_Cnt_F_Ext) & "', '" & SqlCompliantString(m_Cnt_F_Fax) & "', " & " '" & SqlCompliantString(m_Cnt_F_Tel) & "', '', '" & SqlCompliantString(m_Cnt_Email) & "', '" & SqlCompliantString(m_Cnt_Acc_Man) & "', " & " '100', '1', '1', '0', '0', '0', '0', '0' ) "

            db.Execute(strSql)

            strSql = " SELECT ID " & " FROM Addresses " & " WHERE Account = '" & m_Customer.cmp_wwn & "'" & " AND ContactPerson = '" & cnt_id & "' " & " ORDER BY ID DESC "
            db.Execute(strSql)
            dtAddressInfo = db.DataTable(strSql)

            strSql = " INSERT INTO ShippingDays ( AddressID ) " & " VALUES ( '" & dtAddressInfo.Rows(0).Item("ID").ToString & "' ) "
            db.Execute(strSql)

            strSql = " INSERT INTO BacoDiscussionStdPictures ( Description, Filename, Type, Width, Height, Size, Division ) " & " VALUES ( '_" & SqlCompliantString(m_Customer.cmp_code) & "', '" & SqlCompliantString(m_FullName) & "', 'vCard', " & " '1', '1', '1', '100' ) "
            db.Execute(strSql)

        Catch ex As Exception
            MsgBox("Error in CContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Sub SetGender()

        Dim strSql As String
        'Dim rs As New ADODB.Recordset

        Try
            strSql =
            "SELECT     Gender " &
            "FROM       Pred WITH (nolock) " &
            "WHERE      PredCode = '" & m_PredCode & "' "

            'rs.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            'If rs.RecordCount <> 0 Then
            '    m_Gender = rs.Fields("Gender").Value
            'End If
            'rs.Close()

            Dim dt As DataTable
            Dim db As New cDBA

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then
                m_Gender = dt.Rows(0).Item("Gender")
            End If

        Catch er As Exception
            MsgBox("Error in CContact." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Shared Sub contactInstructions(ByVal datafield As String, Optional ByVal orderType As String = "")
        Dim db As New cDBA
        Dim dt As DataTable

        If datafield Is Nothing Then Exit Sub

        Dim subStrsql As String = ""
        If orderType <> "" Then
            subStrsql = " AND ORDER_TYPE = '" & orderType & "' "
        End If

        Dim strsqlMain As String =
           "SELECT     o.DATA_FIELD, ISNULL(o.SPEC_INSTRUCTION, '') AS SPEC_INSTRUCTION, o.SHOW_MESSAGEBOX, o.ORDER_TYPE " &
           "FROM       OEI_ITEM_SPEC_INSTRUCTION as o WITH (NOLOCK) " &
           "WHERE      o.CUS_NO = '" & m_oOrder.Ordhead.Cus_No.Replace("'", "''") & "' AND " &
           "           DATA_FIELD = '" & datafield & "' AND SHOW_MESSAGEBOX = 1 " & subStrsql &
           "ORDER BY   ITEM_NO, DATA_FIELD "

        dt = db.DataTable(strsqlMain)
        If dt.Rows.Count <> 0 Then

            For Each dtRow As DataRow In dt.Rows
                'show popup
                showAppendMessageLog(datafield, dtRow.Item("SPEC_INSTRUCTION").ToString, True)
            Next

        End If
    End Sub

    Private Shared Sub showAppendMessageLog(key As String, msgText As String, Optional ByVal forceShowMessage As Boolean = False)
        Dim isShownAlready = False

        Try
            'check if message was already shown
            If popMessageLog.ContainsKey(key) AndAlso popMessageLog(key).Contains(msgText.Trim()) Then
                isShownAlready = True
            End If

            'update and show message
            If Not isShownAlready Or forceShowMessage = True Then
                If popMessageLog.ContainsKey(key) Then
                    popMessageLog(key) = popMessageLog(key) & "\n" & msgText.Trim()
                Else
                    popMessageLog.Add(key, msgText)
                End If


                MsgBox(Replace(msgText, "\n", vbCrLf))
            End If
        Catch er As Exception
            MsgBox("That was weird from the contacts..." & er.Message)
        End Try
    End Sub

#Region "Public properties ################################################"

    Public Property Cnt_Acc_Man() As String
        Get
            Cnt_Acc_Man = m_Cnt_Acc_Man
        End Get
        Set(ByVal value As String)
            m_Cnt_Acc_Man = value
        End Set
    End Property
    Public Property Cnt_Email() As String
        Get
            Cnt_Email = m_Cnt_Email
        End Get
        Set(ByVal value As String)
            m_Cnt_Email = value
        End Set
    End Property
    Public Property Cnt_F_Ext() As String
        Get
            Cnt_F_Ext = m_Cnt_F_Ext
        End Get
        Set(ByVal value As String)
            m_Cnt_F_Ext = value
        End Set
    End Property
    Public Property Cnt_F_Fax() As String
        Get
            Cnt_F_Fax = m_Cnt_F_Fax
        End Get
        Set(ByVal value As String)
            m_Cnt_F_Fax = value
        End Set
    End Property
    Public Property Cnt_F_Tel() As String
        Get
            Cnt_F_Tel = m_Cnt_F_Tel
        End Get
        Set(ByVal value As String)
            m_Cnt_F_Tel = value
        End Set
    End Property
    Public Property Cnt_ID() As String
        Get
            Cnt_ID = m_Cnt_ID
        End Get
        Set(ByVal value As String)
            m_Cnt_ID = value
        End Set
    End Property
    Public Property Customer() As cCustomer
        Get
            Customer = m_Customer
        End Get
        Set(ByVal value As cCustomer)
            m_Customer = value
        End Set
    End Property
    Public Property FirstName() As String
        Get
            FirstName = m_FirstName
        End Get
        Set(ByVal value As String)
            m_FirstName = value
        End Set
    End Property
    Public Property FullName() As String
        Get
            FullName = m_FullName
        End Get
        Set(ByVal value As String)
            m_FullName = value
        End Set
    End Property
    Public Property Gender() As String
        Get
            Gender = m_Gender
        End Get
        Set(ByVal value As String)
            m_Gender = value
        End Set
    End Property
    Public Property ID() As Integer
        Get
            ID = m_ID
        End Get
        Set(ByVal value As Integer)
            If m_ID = 0 Then
                m_ID = value
                m_OrderContact.ID = value
                m_OrderContact.DefMethod = ContactMethodInt(m_OrderContact.LastSendMethod)
            Else
                m_ID = value
            End If
            m_OrderContact.ID = m_ID
        End Set
    End Property
    Public Property LastName() As String
        Get
            LastName = m_LastName
        End Get
        Set(ByVal value As String)
            m_LastName = value
        End Set
    End Property
    Public Property OrderContact() As COrderContact
        Get
            OrderContact = m_OrderContact
        End Get
        Set(ByVal value As COrderContact)
            m_OrderContact = value
        End Set
    End Property
    Public Property PredCode() As String
        Get
            PredCode = m_PredCode
        End Get
        Set(ByVal value As String)
            m_PredCode = value
            Call SetGender()
        End Set
    End Property
    Public Property TaalCode() As String
        Get
            TaalCode = m_TaalCode
        End Get
        Set(ByVal value As String)
            m_TaalCode = value
        End Set
    End Property

#End Region

End Class
