Public Class cOrderContacts

    Private m_Ord_No As String = ""
    Private m_Ord_Guid As String = ""

    Private m_OEContact As cContact
    Private m_Paperproof As cContact
    Private m_OrderAck As cContact
    Private m_Shipping As cContact
    Private m_PreShipping As cContact
    Private m_Reminder90Day As cContact

    Private m_Method As ContactMethod = ContactMethod.None
    Private m_Info As ContactInfo = ContactInfo.None

    Public Sub New()

        Try
            ' Do nothing, every value is set to default here.
            m_OEContact = New cContact
            m_Paperproof = New cContact
            m_OrderAck = New cContact
            m_Shipping = New cContact
            m_PreShipping = New cContact
            m_Reminder90Day = New cContact

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub New(ByVal pstrOrd_Guid As String)

        Try
            'Set only order number
            'm_Ord_No = pstrOrd_No
            m_Ord_Guid = pstrOrd_Guid
            m_OEContact = New cContact(m_Ord_Guid, ContactType.MainContact)
            m_Paperproof = New cContact(m_Ord_Guid, ContactType.Paperproof)
            m_OrderAck = New cContact(m_Ord_Guid, ContactType.OrderAck)
            m_Shipping = New cContact(m_Ord_Guid, ContactType.Shipping)
            m_PreShipping = New cContact(m_Ord_Guid, ContactType.PreShipping)
            m_Reminder90Day = New cContact(m_Ord_Guid, ContactType.Reminder90day)

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SetOrderContacts(ByRef pOrdHead As cOrdHead)

        Try
            'Set only order number
            'm_Ord_No = pstrOrd_No
            Dim strOrd_Guid As String = pOrdHead.Ord_GUID

            m_OEContact = New cContact(strOrd_Guid, ContactType.MainContact)
            m_Paperproof = New cContact(strOrd_Guid, ContactType.Paperproof)
            m_OrderAck = New cContact(strOrd_Guid, ContactType.OrderAck)
            m_Shipping = New cContact(strOrd_Guid, ContactType.Shipping)
            m_PreShipping = New cContact(strOrd_Guid, ContactType.PreShipping)
            m_Reminder90Day = New cContact(strOrd_Guid, ContactType.Reminder90day)

            Dim db As New cDBA
            Dim dt As DataTable
            Dim strSql As String
            strSql = "SELECT * FROM CICNTP WITH (Nolock) WHERE ISNULL(active_y,0) = 1 And CNT_EMAIL = '" & pOrdHead.Email_Address.Trim() & "' "

            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 
            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                m_OEContact.OrderContact.ID = dt.Rows(0).Item("ID")
                m_OEContact.OrderContact.DefMethod = ContactMethod.Email
                m_OEContact.OrderContact.Save()

                m_Paperproof.OrderContact.ID = dt.Rows(0).Item("ID")
                m_Paperproof.OrderContact.DefMethod = ContactMethod.Email
                m_Paperproof.OrderContact.Save()

                m_OrderAck.OrderContact.ID = dt.Rows(0).Item("ID")
                m_OrderAck.OrderContact.DefMethod = ContactMethod.Email
                m_OrderAck.OrderContact.Save()

                m_Shipping.OrderContact.ID = dt.Rows(0).Item("ID")
                m_Shipping.OrderContact.DefMethod = ContactMethod.Email
                m_Shipping.OrderContact.Save()

                m_PreShipping.OrderContact.ID = dt.Rows(0).Item("ID")
                m_PreShipping.OrderContact.DefMethod = ContactMethod.Email
                m_PreShipping.OrderContact.Save()

                m_Reminder90Day.OrderContact.ID = dt.Rows(0).Item("ID")
                m_Reminder90Day.OrderContact.DefMethod = ContactMethod.Email
                m_Reminder90Day.OrderContact.Save()

            Else

                MsgBox("No contact exist for email " & pOrdHead.Email_Address.Trim() & ". Contact list not created.")

            End If

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub New(ByVal pstrOrd_Guid As String, ByVal pstrOrd_No As String)

        Try
            'Set only order number
            m_Ord_Guid = pstrOrd_Guid
            m_Ord_No = pstrOrd_No

            m_OEContact = New cContact(m_Ord_Guid, ContactType.MainContact)
            m_Paperproof = New cContact(m_Ord_Guid, ContactType.Paperproof)
            m_OrderAck = New cContact(m_Ord_Guid, ContactType.OrderAck)
            m_Shipping = New cContact(m_Ord_Guid, ContactType.Shipping)
            m_PreShipping = New cContact(m_Ord_Guid, ContactType.PreShipping)
            m_Reminder90Day = New cContact(m_Ord_Guid, ContactType.Reminder90day)

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Checks if a contact exists for the order
    Public Function Exists(ByVal OrderNo As String) As Boolean

        Exists = False

        Try

            Dim strSql As String

            strSql =
            " SELECT    CON.FullName as Name, O.defMethod As Method, 1 As isPrimary " &
            " FROM      OEI_ORDER_CONTACTS O WITH (Nolock), cicntp CON WITH (Nolock) " &
            " WHERE ISNULL(CON.active_y,0) = 1  And  O.Ord_Guid = '" & m_Ord_Guid & "' "
            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 

            Dim db As New cDBA
            Dim dt As DataTable = db.DataTable(strSql)

            ' contact exists, is it Primary or Secondary ?
            If dt.Rows.Count <> 0 Then Exists = True

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Checks if a specific contact type exists for the order
    Public Function Exists(ByVal pstrOrd_No As String, ByVal pshrContactType As Short) As Boolean

        Exists = False

        Try

            Dim strSql As String

            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            ' Load Contact from DB
            strSql =
            " SELECT    CON.FullName as Name, O.defMethod As Method, O.primaryContact As isPrimary " &
            " FROM      EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW O WITH (Nolock), cicntp CON WITH (Nolock) " &
            " WHERE ISNULL(CON.active_y,0) = 1 And O.OrderNo = '" & pstrOrd_No & "' AND " &
            "  O.ContactType = '" & pshrContactType & "' "

            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 
            Dim db As New cDBA
            Dim dt As DataTable = db.DataTable(strSql)

            ' contact exists, is it Primary or Secondary ?
            If dt.Rows.Count <> 0 Then Exists = True

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Checks if a specific contact exists for the order
    Public Function Exists(ByVal pstrOrd_No As String, ByVal pintContactID As Integer, ByVal pshrContactType As Short) As Boolean

        Exists = False

        Try

            Dim strSql As String

            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            ' Load Contact from DB
            strSql =
            " SELECT    CON.FullName as Name, O.defMethod As Method, O.primaryContact As isPrimary " &
            " FROM      EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW O WITH (Nolock), cicntp CON WITH (Nolock) " &
            " WHERE  ISNULL(CON.active_y,0) = 1 And O.OrderNo = '" & pstrOrd_No & "' AND " &
            "       O.ContactType = '" & pshrContactType & "' AND " &
            "       O.DefContact = '" & pintContactID & "' AND " &
            "       O.DefContact = CON.ID "

            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 

            Dim db As New cDBA
            Dim dt As DataTable = db.DataTable(strSql)

            ' contact exists, is it Primary or Secondary ?
            If dt.Rows.Count <> 0 Then Exists = True

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Checks if the user is a primary user, if not then it is a secondary user
    Public Function IsPrimary(ByRef OrderNo As String, ByRef ContactID As Integer, ByRef cType_Renamed As Short) As Boolean

        IsPrimary = False

        Try

            Dim strSql As String

            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            ' Load Contact from DB
            strSql =
            " SELECT    CON.FullName as Name, O.defMethod As Method, O.primaryContact As isPrimary " &
            " FROM      EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW O WITH (Nolock), cicntp CON WITH (Nolock) " &
            " WHERE   ISNULL(CON.active_y,0) = 1 And  O.OrderNo = '" & OrderNo & "' AND " &
            "      O.ContactType = '" & cType_Renamed & "' AND " &
            "      O.DefContact = '" & ContactID & "' AND " &
            "      O.DefContact = CON.ID "
            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 
            Dim db As New cDBA
            Dim dt As DataTable = db.DataTable(strSql)

            ' contact exists, is it Primary or Secondary ?
            If dt.Rows.Count <> 0 Then

                If dt.Rows(0).Item("IsPrimary") = 1 Then IsPrimary = True

            End If

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function


    Public Function ContactName(ByVal pintContactID As Integer) As String

        ContactName = ""

        Try

            If pintContactID = 0 Then Exit Function

            Dim strSql As String
            strSql = "SELECT * FROM CiCntp WITH (nolock) where ISNULL(active_y,0) = 1 And ID = " & pintContactID

            '++ID 11.13.2019 added criteria for exclude inactive contacts active_y = 1 

            Dim db As New cDBA
            Dim dt As DataTable = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                ContactName = IIf(dt.Rows(0).Item("FullName").Equals(DBNull.Value), "", dt.Rows(0).Item("FullName"))
            End If

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Private Sub RemoveContact(ByVal pContactType As ContactType)

        Dim strSql As String
        Dim cPri As Short
        Dim intIDToDelete As Integer
        Dim answer As Short
        Dim db As New cDBA

        Try
            If m_oOrder.Ordhead.ExportTS <> "" Then Throw New OEException(OEError.Order_Exported_To_Macola)

            cPri = 1 ' Il n'y a que des 1, la liste secondaire pour le moment est enlevée.
            Select Case pContactType

                Case ContactType.MainContact
                    intIDToDelete = m_OEContact.ID

                Case ContactType.OrderAck
                    intIDToDelete = m_OrderAck.ID

                Case ContactType.Paperproof
                    intIDToDelete = m_Paperproof.ID

                Case ContactType.PreShipping
                    intIDToDelete = m_PreShipping.ID

                Case ContactType.Shipping
                    intIDToDelete = m_Shipping.ID

                Case ContactType.Reminder90day
                    intIDToDelete = m_Reminder90Day.ID

            End Select

            answer = MsgBox("Are you sure you want to delete this contact?", MsgBoxStyle.YesNo, "Delete Contact")

            If answer <> MsgBoxResult.Yes Then Exit Sub

            ' Delete Contact
            strSql = _
            "DELETE " & _
            "FROM   EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_NEW_ " & _
            "WHERE  OrderNo = '" & m_Ord_No & "' AND " & _
            "       DefContact = '" & intIDToDelete & "' AND " & _
            "       ContactType = '" & pContactType & "' AND " & _
            "       PrimaryContact = '" & cPri & "' "

            db.Execute(strSql)

            Call AuditTableEntry(UCase(m_oOrder.Ordhead.Cus_No), m_Ord_No, intIDToDelete, "", pContactType, cPri, "d")

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub AuditTableEntry(ByRef CusNo As String, ByRef OrderNo As String, ByRef ContactID As Integer, ByRef Method As String, ByRef ContactType As Short, ByRef primUser As Short, ByRef modType As String)

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        Dim strSql As String
        Dim db As New cDBA

        Try
            strSql = " INSERT INTO dbo.EXACT_TRAVELER_ORDER_CUSTOMER_CONTACTS_AUDIT " & "     ( cusNo, orderNo, defContact, defMethod, contactType, primaryContact, modType, modDatetime, modUser ) " & " VALUES ( '" & UCase(CusNo) & "', '" & gOrderNo & "', '" & ContactID & "', " & "          '" & Method & "', '" & ContactType & "', '" & primUser & "', '" & modType & "', " & "          GETDATE(), '" & gNTUserID & "' ) "

            ' Execute SQL command
            db.Execute(strSql)

        Catch er As Exception
            MsgBox("Error in COrderContacts." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Save()

        If m_oOrder.Ordhead.ExportTS <> "" Then Exit Sub

        If Not m_OEContact.OrderContact Is Nothing Then
            m_OEContact.OrderContact.Contact = m_OEContact
            m_OEContact.OrderContact.Save()
        End If

        If m_oOrder.Ordhead.User_Def_Fld_4.Trim = String.Empty Then
            m_oOrder.Ordhead.User_Def_Fld_4 = m_OEContact.FullName
            m_oOrder.Ordhead.Save()
        End If

        If m_oOrder.Ordhead.User_Def_Fld_4.Trim = String.Empty Then
            m_oOrder.Ordhead.User_Def_Fld_4 = m_OEContact.FullName
            m_oOrder.Ordhead.Save()
        End If

        If Not m_Paperproof.OrderContact Is Nothing Then
            m_Paperproof.OrderContact.Contact = m_Paperproof
            m_Paperproof.OrderContact.Save()
        End If

        If Not m_PreShipping.OrderContact Is Nothing Then
            m_PreShipping.OrderContact.Contact = m_PreShipping
            m_PreShipping.OrderContact.Save()
        End If

        If Not m_Shipping.OrderContact Is Nothing Then
            m_Shipping.OrderContact.Contact = m_Shipping
            m_Shipping.OrderContact.Save()
        End If

        If Not m_OrderAck.OrderContact Is Nothing Then
            m_OrderAck.OrderContact.Contact = m_OrderAck
            m_OrderAck.OrderContact.Save()
        End If

        If Not m_Reminder90Day.OrderContact Is Nothing Then
            m_Reminder90Day.OrderContact.Contact = m_Reminder90Day
            m_Reminder90Day.OrderContact.Save()
        End If

        Call Set_Alternate_Rep()

    End Sub

    Private Sub Set_Alternate_Rep()

        Dim oCicntp As New cCicntp(m_OEContact.OrderContact.ID)

        If oCicntp.Alternate_Rep <> 0 Then
            Dim oC As New cHumres()
            oC.Load_By_Res_ID(oCicntp.Alternate_Rep)
            'm_oOrder.Ordhead.Slspsn_No = oCicntp.Alternate_Rep
            m_oOrder.Ordhead.User_Def_Fld_3 = oC.Fullname
            m_oOrder.Ordhead.Save()
        End If

    End Sub

#Region "Public properties ################################################"

    Public Property OEContact() As cContact
        Get
            OEContact = m_OEContact
        End Get
        Set(ByVal value As cContact)
            m_OEContact = value
        End Set
    End Property

    Public Property OrderAck() As cContact
        Get
            OrderAck = m_OrderAck
        End Get
        Set(ByVal value As cContact)
            m_OrderAck = value
        End Set
    End Property

    Public Property Paperproof() As cContact
        Get
            Paperproof = m_Paperproof
        End Get
        Set(ByVal value As cContact)
            m_Paperproof = value
        End Set
    End Property

    Public Property PreShipping() As cContact
        Get
            PreShipping = m_PreShipping
        End Get
        Set(ByVal value As cContact)
            m_PreShipping = value
        End Set
    End Property

    Public Property Reminder90Day() As cContact
        Get
            Reminder90Day = m_Reminder90Day
        End Get
        Set(ByVal value As cContact)
            m_Reminder90Day = value
        End Set
    End Property

    Public Property Shipping() As cContact
        Get
            Shipping = m_Shipping
        End Get
        Set(ByVal value As cContact)
            m_Shipping = value
        End Set
    End Property

#End Region

End Class
