Public Class cOEI_PRE_READ_ORDER_EXCEPTION

    Private m_ClassName As String = "cOEI_PRE_READ_ORDER_EXCEPTION"
    Private db As New cDBA

    Private m_ID As Integer
    Private m_strOrder_No As String
    Private m_strPo_No As String
    Private m_strItem_Cd As String
    Private m_strItem_No As String
    Private m_intException_TP As Integer
    Private m_strTimestamp As String
    Private m_strEnterPerson As String
    Private m_strUserId As String
    Private m_strUser_Login As String
    Private m_dtCreate_Ts As DateTime
    Private m_dtUpdate_Ts As DateTime
    Private m_strException As String

    'Public Sub New(ByVal in_OrderNo As String, ByVal in_Exception_type As String)
    Public Sub New()
        Id = 0
        ORDER_NO = ""
        OE_PO_NO = ""
        Exception_type = 2793
        USER_LOGIN = g_User.FullName
        ENTER_person = g_User.FullName
        timestamp = ""
        USERID = g_User.Usr_ID
        Create_Ts = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
        Update_Ts = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub

    Public Sub New(ByVal in_OE_PO_NO As String, ByVal in_ORDER_NO As String, ByVal in_item_no As String, ByVal in_Exception_type As Int32)
        Id = 0
        ORDER_NO = in_ORDER_NO
        OE_PO_NO = in_OE_PO_NO
        ITEM_NO = in_item_no
        Exception_type = in_Exception_type
        USER_LOGIN = g_User.FullName
        ENTER_person = g_User.FullName
        timestamp = ""
        USERID = g_User.Usr_ID
        Create_Ts = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
        Update_Ts = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub


    Public Property Id() As Int32
        Get
            Id = m_ID
        End Get
        Set(ByVal value As Int32)
            m_ID = value
        End Set
    End Property

    Public Property ORDER_NO() As String
        Get
            ORDER_NO = m_strOrder_No
        End Get
        Set(ByVal value As String)
            m_strOrder_No = value
        End Set
    End Property

    Public Property OE_PO_NO() As String
        Get
            OE_PO_NO = m_strPo_No
        End Get
        Set(ByVal value As String)
            m_strPo_No = value
        End Set
    End Property


    Public Property ITEM_CD() As String
        Get
            ITEM_CD = m_strItem_Cd
        End Get
        Set(ByVal value As String)
            m_strItem_Cd = value
        End Set
    End Property

    Public Property ITEM_NO() As String
        Get
            ITEM_NO = m_strItem_No
        End Get
        Set(ByVal value As String)
            m_strItem_No = value
        End Set
    End Property

    Public Property Exception_type() As Int32
        Get
            Exception_type = m_intException_TP
        End Get
        Set(ByVal value As Int32)
            m_intException_TP = value
        End Set
    End Property

    Public Property timestamp() As String
        Get
            timestamp = m_strTimestamp
        End Get
        Set(ByVal value As String)
            m_strTimestamp = value
        End Set
    End Property

    Public Property ENTER_person() As String
        Get
            ENTER_person = m_strEnterPerson
        End Get
        Set(ByVal value As String)
            m_strEnterPerson = value
        End Set
    End Property

    Public Property USERID() As String
        Get
            USERID = m_strUserId
        End Get
        Set(ByVal value As String)
            m_strUserId = value
        End Set
    End Property

    Public Property USER_LOGIN() As String
        Get
            USER_LOGIN = m_strUser_Login
        End Get
        Set(ByVal value As String)
            m_strUser_Login = value
        End Set
    End Property

    Public Property Create_Ts() As DateTime
        Get
            Create_Ts = m_dtCreate_Ts
        End Get
        Set(ByVal value As DateTime)
            m_dtCreate_Ts = value
        End Set
    End Property

    Public Property Update_Ts() As DateTime
        Get
            Update_Ts = m_dtUpdate_Ts
        End Get
        Set(ByVal value As DateTime)
            m_dtUpdate_Ts = value
        End Set
    End Property


    Public Property Exception_text() As String
        Get
            Exception_text = m_strException
        End Get
        Set(ByVal value As String)
            m_strException = value
        End Set
    End Property
End Class
