Public Class cMdb_Cfg_Enum


#Region "private variables #################################################"


    Private m_intId As Int32
    Private m_strEnum_Cat As String
    Private m_strEnum_Value As String
    Private m_strEnum_Sub_Cat As String
    Private m_strUser_Login As String
    Private m_dtCreate_Ts As DateTime
    Private m_dtUpdate_Ts As DateTime

#End Region


#Region "Public properties ################################################"
    Public Sub New()
        Id = 0
        Enum_Cat = ""
        Enum_Sub_Cat = ""
        Enum_Value = ""
        Update_Ts = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub


    Public Sub New(ByVal id_ As Integer, ByVal name As String)
        Id = id_
        Enum_Cat = ""
        Enum_Sub_Cat = ""
        Enum_Value = ""
        Update_Ts = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub

    Public Sub New(ByVal cat As String)
        Id = 0
        Enum_Cat = ""
        Enum_Sub_Cat = cat
        Enum_Value = ""
        Update_Ts = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub

    Public Sub New(ByVal dr As DataGridViewRow)
        If dr.Cells.Item("Id").Value IsNot DBNull.Value Then Id = dr.Cells.Item("Id").Value
        If dr.Cells.Item("Enum_Cat").Value IsNot DBNull.Value Then Enum_Cat = dr.Cells.Item("Enum_Cat").Value
        If dr.Cells.Item("Enum_Sub_Cat").Value IsNot System.DBNull.Value Then Enum_Sub_Cat = dr.Cells.Item("Enum_Sub_Cat").Value
        If dr.Cells.Item("Enum_Value").Value IsNot DBNull.Value Then Enum_Value = dr.Cells.Item("Enum_Value").Value
        Update_Ts = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub

    Public Property Id() As Int32
        Get
            Id = m_intId
        End Get
        Set(ByVal value As Int32)
            m_intId = value
        End Set
    End Property

    Public Property Enum_Cat() As String
        Get
            Enum_Cat = m_strEnum_Cat
        End Get
        Set(ByVal value As String)
            m_strEnum_Cat = value
        End Set
    End Property



    Public Property Enum_Value() As String
        Get
            Enum_Value = m_strEnum_Value
        End Get
        Set(ByVal value As String)
            m_strEnum_Value = value
        End Set
    End Property

    Public Property Enum_Sub_Cat() As String
        Get
            Enum_Sub_Cat = m_strEnum_Sub_Cat
        End Get
        Set(ByVal value As String)
            m_strEnum_Sub_Cat = value
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

    Public Property User_Login() As String
        Get
            User_Login = m_strUser_Login
        End Get
        Set(ByVal value As String)
            m_strUser_Login = value
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

#End Region



End Class ' cMdb_Cfg_Enum


