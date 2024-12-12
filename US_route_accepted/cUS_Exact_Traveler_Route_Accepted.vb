Public Class cUS_Exact_Traveler_Route_Accepted


#Region "private variables #################################################"

    Private m_US_Route_Accepted_ID As Int32
    Private m_CAT As String
    Private m_RouteCatID As Int32
    Private m_RouteCategory As String
    Private m_deco_meth_1 As String
    Private m_deco_meth_2 As String
    Private m_RouteID As Int32
    Private m_RouteDescription As String


#End Region


#Region "Public properties ################################################"
    Public Sub New()

        m_CAT = ""
        m_RouteCatID = 0
        m_RouteCategory = ""
        m_deco_meth_1 = ""
        m_deco_meth_2 = ""
        m_RouteID = 0
        m_RouteDescription = ""
    End Sub
    Public Property US_Route_Accepted_ID() As Int32
        Get
            US_Route_Accepted_ID = m_US_Route_Accepted_ID
        End Get
        Set(ByVal value As Int32)
            m_US_Route_Accepted_ID = value
        End Set
    End Property

    Public Property CAT() As String
        Get
            CAT = m_CAT
        End Get
        Set(ByVal value As String)
            m_CAT = value
        End Set
    End Property

    Public Property RouteCatID() As Int32
        Get
            RouteCatID = m_RouteCatID
        End Get
        Set(ByVal value As Int32)
            m_RouteCatID = value
        End Set
    End Property

    Public Property RouteCategory() As String
        Get
            RouteCategory = m_RouteCategory
        End Get
        Set(ByVal value As String)
            m_RouteCategory = value
        End Set
    End Property

    Public Property deco_meth_1() As String
        Get
            deco_meth_1 = m_deco_meth_1
        End Get
        Set(ByVal value As String)
            m_deco_meth_1 = value
        End Set
    End Property

    Public Property deco_meth_2() As String
        Get
            deco_meth_2 = m_deco_meth_2
        End Get
        Set(ByVal value As String)
            m_deco_meth_2 = value
        End Set
    End Property

    Public Property RouteID() As Int32
        Get
            RouteID = m_RouteID
        End Get
        Set(ByVal value As Int32)
            m_RouteID = value
        End Set
    End Property

    Public Property RouteDescription() As String
        Get
            RouteDescription = m_RouteDescription
        End Get
        Set(ByVal value As String)
            m_RouteDescription = value
        End Set
    End Property

#End Region


End Class ' cUS_Exact_Traveler_Route_Accepted


