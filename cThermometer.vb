Public Class cThermometer

    Private m_ThermoStyle As StyleEnum = StyleEnum.Full
    Private m_MethodName As String = String.Empty
    Private m_RouteID As Integer = 0
    Private m_MethodCollection As New Collection

    Private db As New cDBA()
    Private dt As DataTable
    Private strSql As String
    Private dtRow As DataRow

    Public Sub New()

        ' Default: Set m_MethodCollection to a collection containing every method
        Try

            m_ThermoStyle = StyleEnum.Full
            Call LoadAllMethods()

        Catch er As Exception
            MsgBox("Error in CThermometer." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub New(ByVal oStyle As StyleEnum)

        ' Set m_MethodCollection to a collection containing every method
        Try

            m_ThermoStyle = oStyle
            Call LoadAllMethods()

        Catch er As Exception
            MsgBox("Error in CThermometer." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub New(ByVal oStyle As StyleEnum, ByVal pMethodName As String)

        ' Set m_MethodCollection to a collection containing only the item pMethodName
        Try

            m_ThermoStyle = oStyle
            m_MethodName = pMethodName

            Call LoadMethodsFromName()

        Catch er As Exception
            MsgBox("Error in CThermometer." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub New(ByVal oStyle As StyleEnum, ByVal pRouteID As Integer)

        ' Set m_MethodCollection to a collection containing every method of the routecategory for RouteID
        Try

            m_ThermoStyle = oStyle
            m_RouteID = pRouteID

            Call LoadMethodsFromRouteID()

        Catch er As Exception
            MsgBox("Error in CThermometer." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadAllMethods()

        Try

            strSql = _
            "SELECT DISTINCT    Method_Name " & _
            "FROM               EXACT_TRAVELER_THERMO_AVAIL_DATE WITH (Nolock) "

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                For Each dtRow In dt.Rows
                    m_MethodCollection.Add(dtRow.Item("Method_Name").ToString, dtRow.Item("Method_Name").ToString)
                Next dtRow
            End If

        Catch er As Exception
            MsgBox("Error in CThermometer." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadMethodsFromRouteID()

        Try

            Dim routeEl As String ' Object

            strSql = _
            "SELECT ISNULL(RouteCategory, '') AS RouteCategory " & _
            "FROM   EXACT_TRAVELER_ROUTE " & _
            "WHERE  RouteID = " & m_RouteID.ToString

            dt = db.DataTable(strSql)
            If dt.Rows.Count <> 0 Then

                Dim arrRoutes() As String
                arrRoutes = dt.Rows(0).Item("RouteCategory").ToString.Split("/")

                For Each routeEl In arrRoutes

                    m_MethodCollection.Add(routeEl, routeEl)

                Next

            End If

        Catch er As Exception
            MsgBox("Error in CThermometer." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadMethodsFromName()

        Try

            m_MethodCollection.Add(m_MethodName, m_MethodName)

        Catch er As Exception
            MsgBox("Error in CThermometer." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Show()

        Try
            Dim ofrmThermometer As New frmThermometer(Me)
            ofrmThermometer.ShowDialog()
            ofrmThermometer.Dispose()

        Catch er As Exception
            MsgBox("Error in CThermometer." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Property Style() As StyleEnum
        Get
            Style = m_ThermoStyle
        End Get
        Set(ByVal value As StyleEnum)
            m_ThermoStyle = value
        End Set
    End Property

    Public Property Method_Name() As String
        Get
            Method_Name = m_MethodName
        End Get
        Set(ByVal value As String)
            m_MethodName = value
        End Set
    End Property

    Public Property Methods() As Collection
        Get
            Methods = m_MethodCollection
        End Get
        Set(ByVal value As Collection)
            m_MethodCollection = value
        End Set
    End Property
    Public Enum StyleEnum
        Full
        Method
        Route
    End Enum

End Class
