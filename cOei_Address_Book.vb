Public Class cOei_Address_Book


#Region "private variables #################################################" 


    Private m_Id As Int32
    Private m_Fullname As String
    Private m_Mail As String 


#End Region


#Region "Public properties ################################################" 
    Public Property addrBook_Id() As Int32
        Get
            addrBook_Id = m_Id
        End Get
        Set(ByVal value As Int32)
            m_Id = value
        End Set
    End Property

    Public Property addrBook_Fullname() As String
        Get
            addrBook_Fullname = m_Fullname
        End Get
        Set(ByVal value As String)
            m_Fullname = value
        End Set
    End Property

    Public Property addrBook_Mail() As String
        Get
            addrBook_Mail = m_Mail
        End Get
        Set(ByVal value As String)
            m_Mail = value
        End Set
    End Property
#End Region 


End Class ' cOei_Address_Book


