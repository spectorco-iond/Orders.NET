﻿Public Class cUser

    Private m_ID As Integer = 0
    Private m_HumRes_ID As Integer = 0
    Private m_Usr_ID As String = ""
    Private m_Blocked As Integer = 0
    Private m_First_Name As String = ""
    Private m_Sur_Name As String = ""
    Private m_FullName As String = ""
    Private m_Name_Prefix As String = ""
    Private m_Order_Count As Integer = 0
    Private m_Lock_Code As String = ""
    Private m_Mail As String = ""
    Private m_Group_Code As String = ""
    Private m_Group_Defaults As New Collection ' So Item count is 0
    Private m_Conn_String As String = ""
    Private m_EDI_Access As String = 0
    Private m_OEI_Version As String = ""
    Private m_Last_Login_TS As DateTime
    Private m_Xml_Access As Integer = 0

#Region "Public Constructors ##############################################"

    ' On ne met rien dans ce constructeur, pour eviter de travailler chaque fois qu'on 
    ' transporte l'object dans une autre classe.
    Public Sub New()

        'm_Usr_ID = GetNTUserID()
        'Call Load()

    End Sub

    Public Sub New(ByVal pstrUsr_ID As String, ByVal pstrConn_String As String)

        m_Conn_String = pstrConn_String

        If pstrUsr_ID = "" Then
            m_Usr_ID = GetNTUserID()
        Else
            m_Usr_ID = pstrUsr_ID
        End If

        Call Load()

    End Sub

    Public Sub New(ByVal pstrUsr_ID As String)

        If pstrUsr_ID = "" Then
            m_Usr_ID = GetNTUserID()
        Else
            m_Usr_ID = pstrUsr_ID
        End If

        Call Load()

    End Sub

    Public Sub New(ByVal pintID As Integer)

        m_ID = pintID
        Call Load()

    End Sub

#End Region

#Region "Public I/O routines ##############################################"

    'Public Function Exists() As Boolean

    '    Dim strSql As String
    '    Dim conn As New cDBA("Exact")
    '    Dim dtHumRes As New DataTable

    '    strSql = _
    '    "SELECT * " & _
    '    "FROM   OEI_Users WITH (Nolock) " & _
    '    "WHERE  HumRes_ID = " & m_HumRes_ID & " "

    '    dtHumRes = New DataTable
    '    dtHumRes = conn.DataTable(strSql)

    '    Exists = (dtHumRes.Rows.Count <> 0)

    'End Function

    Private Function Exists(ByVal pstrPrefix As String) As Boolean

        Exists = False

        Dim strSql As String
        Dim conn As New cDBA
        Dim dtHumRes As New DataTable

        strSql =
        "SELECT * " &
        "FROM   OEI_Users WITH (Nolock) " &
        "WHERE  Name_Prefix = '" & pstrPrefix.ToUpper & "' "

        dtHumRes = New DataTable
        dtHumRes = conn.DataTable(strSql)

        Exists = (dtHumRes.Rows.Count <> 0)

    End Function

    Private Sub Load()

        Call LoadFromHumRes()
        Call Save()

    End Sub

    Private Sub LoadFromHumRes()

        Dim strSql As String
        'Dim conn As New cDBA
        Dim conn As cDBA

        Dim dtHumRes As DataTable

        If m_Conn_String = "" Then
            conn = New cDBA(g_Default_Conn_String)
        Else
            conn = New cDBA(m_Conn_String)
        End If

        Try
            If m_ID = 0 Then
                strSql =
                "SELECT * " &
                "FROM  HumRes WITH (Nolock) " &
                "WHERE  Usr_ID = '" & m_Usr_ID & "' "
            Else
                strSql = _
                "SELECT * " & _
                "FROM   HumRes WITH (Nolock) " & _
                "WHERE  ID = '" & m_ID & "' "
            End If

            'MsgBox(strSql & " --- " & m_ID)

            dtHumRes = conn.DataTable(strSql)
            If dtHumRes.Rows.Count = 0 Then
                Throw New OEException(OEError.User_Does_Not_Exist_In_Macola)
            Else
                m_HumRes_ID = dtHumRes.Rows(0).Item("ID")
                m_Usr_ID = dtHumRes.Rows(0).Item("Usr_ID")
                m_Blocked = dtHumRes.Rows(0).Item("Blocked")
                m_First_Name = dtHumRes.Rows(0).Item("First_Name").ToString
                m_Sur_Name = dtHumRes.Rows(0).Item("Sur_Name").ToString
                m_FullName = dtHumRes.Rows(0).Item("FullName").ToString
                m_Mail = dtHumRes.Rows(0).Item("Mail").ToString
                'If m_Blocked = 0 Then Throw New OEException(OEError.User_Is_Blocked_From_Macola)
            End If

        Catch er As Exception
            MsgBox("Error in CUser." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Function NextOrderNumber() As String
        '==========
        'The NEW order numbers for each user are determined in part from the user's name.
        'Eg.,  User HVI07 creates orders that begin with "HV" -  HV00001 and increment from there.
        'This looks up the highest order number that matches the prefix and returns the next number to be used.
        '==========
        NextOrderNumber = ""

        Dim strMaxOrderNumber As String
        Dim intNextOrderNumber As Integer

        Dim strSql As String
        Dim db As New cDBA()
        Dim db2 As New cDBA()

        Dim dtHumRes As New DataTable

        strSql =
        "SELECT * " &
        "FROM   OEI_Users " &
        "WHERE  HumRes_ID = '" & m_HumRes_ID & "' "

        dtHumRes = New DataTable
        dtHumRes = db.DataTable(strSql)

        If dtHumRes.Rows.Count <> 0 Then

            With dtHumRes.Rows(0)
                'Find maximum order number that starts with the users prefix to allow other users to submit another user's order
                'without screwing up the next order number generated sequence.

                strSql =
               "SELECT		U.Lock_Code, U.Order_Count, MAX(ISNULL(H.OEI_Ord_No,RTRIM(UPPER(U.Name_Prefix))  + '0')) AS OEI_Ord_No " &
               "FROM		OEI_USERS U WITH (Nolock)  " &
               "LEFT JOIN	OEI_ORDHDR H WITH (Nolock) ON U.Name_Prefix = substring(OEI_Ord_No,1,2) " &
               "WHERE		U.HumRes_ID = " & m_HumRes_ID & " " &
               "GROUP BY	U.Lock_Code, U.Order_Count "

            End With
            'MsgBox("02" & strSql)

            ' We must put dtOrder under db2 or the update will not be valid.
            Dim dtOrder As DataTable = db2.DataTable(strSql)

            If dtOrder.Rows.Count > 0 Then ' dtOrder.Rows.Count <> 0 Then
                strMaxOrderNumber = Trim(dtOrder.Rows(0).Item("OEI_Ord_No"))
            Else
                strMaxOrderNumber = "0" ' This will not occur
            End If

            'update this expression to increment the count based on users order_count number
            intNextOrderNumber = CInt(strMaxOrderNumber.Replace(Trim(dtHumRes.Rows(0).Item("Name_Prefix")), "")) + 1 'CInt(Trim(dtHumRes.Rows(0).Item("Order_Count")))
            dtHumRes.Rows(0).Item("Order_Count") = intNextOrderNumber

            Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
            db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            db.DBDataAdapter.Update(db.DBDataTable)

            m_Order_Count = intNextOrderNumber
            NextOrderNumber = Trim(m_Name_Prefix).ToString & CStr(intNextOrderNumber).ToString.PadLeft(6, "0"c)

        End If

    End Function


    Public Function OrderNumberGenerated() As String

        OrderNumberGenerated = ""

        Dim intOrder_Count As Integer

        Dim strSql As String
        Dim conn As New cDBA()
        Dim dtHumRes As New DataTable

        strSql = _
        "SELECT ISNULL(Order_Count, 0) AS Order_Count " & _
        "FROM   OEI_Users " & _
        "WHERE  HumRes_ID = '" & m_HumRes_ID & "' "

        dtHumRes = New DataTable
        dtHumRes = conn.DataTable(strSql)

        If dtHumRes.Rows.Count <> 0 Then
            intOrder_Count = dtHumRes.Rows(0).Item("Order_Count")
            OrderNumberGenerated = Trim(m_Name_Prefix).ToString & CStr(intOrder_Count).ToString.PadLeft(6, "0"c)
        End If

    End Function

    Public Function PreviousOrderNumber() As String

        PreviousOrderNumber = ""

        Dim intPreviousOrderNumber As Integer

        Dim strSql As String
        Dim db As New cDBA()
        Dim dtHumRes As DataTable

        Dim dtOrder As DataTable
        strSql = "SELECT * FROM OEI_ORDHDR WITH (NOLOCK) WHERE OEI_Ord_No = '" & OrderNumberGenerated() & "' "
        dtOrder = db.DataTable(strSql)
        If dtOrder.Rows.Count <> 0 Then
            PreviousOrderNumber = OrderNumberGenerated()
            Exit Function
        End If

        strSql = _
        "SELECT * " & _
        "FROM   OEI_Users " & _
        "WHERE  HumRes_ID = '" & m_HumRes_ID & "' "

        dtHumRes = db.DataTable(strSql)

        If dtHumRes.Rows.Count <> 0 Then

            intPreviousOrderNumber = dtHumRes.Rows(0).Item("Order_Count") - 1
            If intPreviousOrderNumber < 0 Then
                intPreviousOrderNumber = 0
            End If

            dtHumRes.Rows(0).Item("Order_Count") = intPreviousOrderNumber

            m_Order_Count = intPreviousOrderNumber

            Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
            db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            db.DBDataAdapter.Update(db.DBDataTable)

            'If Trim(m_Name_Prefix).Length = 2 Then
            '    m_Name_Prefix = Trim(m_Name_Prefix) & "0"
            'End If

            PreviousOrderNumber = Trim(m_Name_Prefix).ToString & CStr(intPreviousOrderNumber).ToString.PadLeft(6, "0"c)
            'NextOrderNumber = m_Name_Prefix

        End If

    End Function

    Public Sub Save()

        Dim strSql As String
        Dim db As New cDBA()
        Dim dtHumRes As New DataTable


        '++ID  11.14.2019 
        '  m_OEI_Version = "1.0.0.173" ' Application.ProductVersion.ToString"

        Dim _assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly
        m_OEI_Version = _assembly.GetName.Version.ToString()

        strSql = _
        "SELECT * " & _
        "FROM   OEI_Users WITH (Nolock) " & _
        "WHERE  HumRes_ID = '" & m_HumRes_ID & "' "

        dtHumRes = New DataTable
        dtHumRes = db.DataTable(strSql)
        Dim drDataRow As DataRow

        If dtHumRes.Rows.Count <> 0 Then

            drDataRow = dtHumRes.Rows(0)
            m_ID = drDataRow.Item("ID")
            m_Name_Prefix = drDataRow.Item("Name_Prefix")
            m_Order_Count = drDataRow.Item("Order_Count")
            m_Lock_Code = drDataRow.Item("Lock_Code")
            m_Group_Code = IIf(drDataRow.Item("Group_Code").Equals(DBNull.Value), "", drDataRow.Item("Group_Code").ToString)
            m_Conn_String = drDataRow.Item("Conn_String")
            m_EDI_Access = drDataRow.Item("EDI_Access")
            m_Xml_Access = drDataRow.Item("Xml_Access")

            'Dim AD As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
            'm_OEI_Version = AD.CurrentVersion.ToString

            If m_Group_Code <> "" Then

                Dim dtGroup_Defaults As DataTable

                strSql = _
                "SELECT     * " & _
                "FROM       OEI_Group_Defaults WITH (Nolock) " & _
                "WHERE      Group_Code = '" & m_Group_Code & "' " & _
                "ORDER BY   Field_Name "

                dtGroup_Defaults = db.DataTable(strSql)

                If dtGroup_Defaults.Rows.Count <> 0 Then

                    For Each drRow As DataRow In dtGroup_Defaults.Rows
                        If Not (m_Group_Defaults.Contains(Trim(drRow.Item("Field_Name").ToString))) Then
                            m_Group_Defaults.Add(Trim(drRow.Item("Default_Value").ToString), Trim(drRow.Item("Field_Name").ToString))
                        End If
                    Next

                End If

            End If

            'm_OEI_Version = "1.0.0.161" ' Application.ProductVersion.ToString"
            drDataRow.Item("OEI_Version") = m_OEI_Version

            m_Last_Login_TS = Date.Now
            drDataRow.Item("Last_Login_TS") = m_Last_Login_TS

            Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
            db.DBDataAdapter.UpdateCommand = cmd.GetUpdateCommand
            db.DBDataAdapter.Update(db.DBDataTable)

        Else

            drDataRow = dtHumRes.NewRow()
            drDataRow.Item("HumRes_ID") = m_HumRes_ID
            'drDataRow.Item("Name_Prefix") = (Mid(m_First_Name, 1, 1) & Mid(m_Sur_Name, 1)).ToUpper
            drDataRow.Item("Conn_String") = m_Conn_String
            drDataRow.Item("EDI_Access") = m_EDI_Access
            drDataRow.Item("XML_Access") = m_Xml_Access

            Dim intCounter As Integer = 0
            Dim pstrPrefix As String = Mid(Trim(m_First_Name), 1, 1) & Mid(Trim(m_Sur_Name), 1, 1).ToUpper
            If pstrPrefix.Length = 1 Then pstrPrefix &= "A"

            While Exists(IIf(intCounter = 0, pstrPrefix, pstrPrefix & Chr(64 + intCounter)))
                intCounter += 1
            End While

            If intCounter > 0 Then pstrPrefix = pstrPrefix & Chr(64 + intCounter)
            drDataRow.Item("Name_Prefix") = pstrPrefix
            drDataRow.Item("Order_Count") = 0 ' m_Order_Count
            drDataRow.Item("Lock_Code") = m_Usr_ID
            m_Name_Prefix = pstrPrefix

            'm_OEI_Version = Application.ProductVersion.ToString
            drDataRow.Item("OEI_Version") = m_OEI_Version

            m_Last_Login_TS = Date.Now
            drDataRow.Item("Last_Login_TS") = m_Last_Login_TS

            db.DBDataTable.Rows.Add(drDataRow)
            Dim cmd As Odbc.OdbcCommandBuilder = New Odbc.OdbcCommandBuilder(db.DBDataAdapter)
            db.DBDataAdapter.InsertCommand = cmd.GetInsertCommand
            db.DBDataAdapter.Update(db.DBDataTable)

        End If

    End Sub

#End Region

#Region "Public properties ################################################"

    ' User is blocked when value is 0, unblocked if value is 1. From HumRes.Blocked.
    Public Property Blocked() As Integer
        Get
            Blocked = m_Blocked
        End Get
        Set(ByVal value As Integer)
            m_Blocked = value
        End Set
    End Property
    Public Property Conn_String() As String
        Get
            Conn_String = m_Conn_String
        End Get
        Set(ByVal value As String)
            m_Conn_String = value
        End Set
    End Property
    Public Property EDI_Access() As Integer
        Get
            EDI_Access = m_EDI_Access
        End Get
        Set(ByVal value As Integer)
            m_EDI_Access = value
        End Set
    End Property
    Public Property First_Name() As String
        Get
            First_Name = m_First_Name
        End Get
        Set(ByVal value As String)
            m_First_Name = value
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
    Public Property Group_Code() As String
        Get
            Group_Code = m_Group_Code
        End Get
        Set(ByVal value As String)
            m_Group_Code = value
        End Set
    End Property
    Public ReadOnly Property Group_Defaults() As Collection
        Get
            Group_Defaults = m_Group_Defaults
        End Get
    End Property
    Public Property HumRes_ID() As Integer
        Get
            HumRes_ID = m_HumRes_ID
        End Get
        Set(ByVal value As Integer)
            m_HumRes_ID = value
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
    Public Property Lock_Code() As String
        Get
            Lock_Code = m_Lock_Code
        End Get
        Set(ByVal value As String)
            m_Lock_Code = value
        End Set
    End Property
    Public Property Mail() As String
        Get
            Mail = m_Mail
        End Get
        Set(ByVal value As String)
            m_Mail = value
        End Set
    End Property
    Public Property Name_Prefix() As String
        Get
            Name_Prefix = m_Name_Prefix
        End Get
        Set(ByVal value As String)
            m_Name_Prefix = value
        End Set
    End Property
    Public Property Order_Count() As Integer
        Get
            Order_Count = m_Order_Count
        End Get
        Set(ByVal value As Integer)
            m_Order_Count = value
        End Set
    End Property
    Public Property Sur_Name() As String
        Get
            Sur_Name = m_Sur_Name
        End Get
        Set(ByVal value As String)
            m_Sur_Name = value
        End Set
    End Property
    Public Property Usr_ID() As String
        Get
            Usr_ID = m_Usr_ID
        End Get
        Set(ByVal value As String)
            m_Usr_ID = value
        End Set
    End Property
    Public Property XML_Access() As Integer
        Get
            XML_Access = m_Xml_Access
        End Get
        Set(ByVal value As Integer)
            m_Xml_Access = value
        End Set
    End Property

#End Region

End Class
