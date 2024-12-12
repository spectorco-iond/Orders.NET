Imports System.Data.Odbc

Public Class cDBA

    Private m_Connexion As Odbc.OdbcConnection
    Private m_Command As Odbc.OdbcCommand
    Private m_CommandBuilder As Odbc.OdbcCommandBuilder
    Private m_DataAdapter As Odbc.OdbcDataAdapter
    Private m_DataSet As DataSet
    Private m_DataTable As DataTable

    Public Sub New()



        m_Connexion = Connexion() ' New Odbc.OdbcConnection
        m_Command = New Odbc.OdbcCommand
        m_CommandBuilder = New Odbc.OdbcCommandBuilder
        m_DataAdapter = New Odbc.OdbcDataAdapter
        m_DataSet = New DataSet
        m_DataTable = New DataTable

    End Sub

    Public Sub New(ByVal pstrConnection As String)

        m_Connexion = Connexion(pstrConnection) ' New Odbc.OdbcConnection
        m_Command = New Odbc.OdbcCommand
        m_CommandBuilder = New Odbc.OdbcCommandBuilder
        m_DataAdapter = New Odbc.OdbcDataAdapter
        m_DataSet = New DataSet
        m_DataTable = New DataTable

    End Sub

    Public Function Connexion() As Odbc.OdbcConnection

        'Connexion = Connexion("Driver={SQL Server};Server=Exact;Trusted_Connection=Yes;Database=100")
        'Connexion = Connexion("Provider=SQLOLEDB;" & gConn.ConnectionString)
        '==============================
        'CHANGED Aug 3, 2017 - T. Louzon
        If gDSNName Is Nothing Then
            Connexion = Connexion(g_Default_Conn_String)
        Else
            Connexion = Connexion("DSN=" & gDSNName)
        End If

        '==============================
        'CHANGED Aug 3, 2017 - T. Louzon
        'If gDSNName Is Nothing Then
        '    If g_User Is Nothing Then
        '        Connexion = Connexion(g_Default_Conn_String)
        '    Else
        '        Connexion = Connexion(g_User.Conn_String)
        '    End If
        'Else
        '    If gDSNName.ToUpper.Trim = "THOR" Then
        '        Connexion = Connexion("DSN=Thor;")
        '    Else
        '        If g_User Is Nothing Then
        '            Connexion = Connexion(g_Default_Conn_String)
        '        Else
        '            Connexion = Connexion(g_User.Conn_String)
        '        End If
        '    End If
        'End If
        '==============================

        m_Connexion = Connexion

        'ADDED July 14, 2017
        'Test Connection and get databasename/servername for login screen
        'Connection must be open to set the value
        Connexion.Open()
        If Connexion.Database.ToString <> "" Then gsDataBaseName = Connexion.Database.ToString
        If Connexion.DataSource.ToString <> "" Then gsServerName = Connexion.DataSource.ToString
        Connexion.Close()

    End Function

    Public Property DBCommand() As Odbc.OdbcCommand
        Get
            DBCommand = m_Command
        End Get
        Set(ByVal value As Odbc.OdbcCommand)
            m_Command = value
        End Set
    End Property

    Public Property DBDataAdapter() As Odbc.OdbcDataAdapter
        Get
            DBDataAdapter = m_DataAdapter
        End Get
        Set(ByVal value As Odbc.OdbcDataAdapter)
            m_DataAdapter = value
        End Set
    End Property

    Public Property DBDataSet() As DataSet
        Get
            DBDataSet = m_DataSet
        End Get
        Set(ByVal value As DataSet)
            m_DataSet = value
        End Set
    End Property

    Public Property DBDataTable() As DataTable
        Get
            DBDataTable = m_DataTable
        End Get
        Set(ByVal value As DataTable)
            m_DataTable = value
        End Set
    End Property

    Public Function Connexion(ByVal pstrConnectionString) As Odbc.OdbcConnection

        Connexion = New Odbc.OdbcConnection(pstrConnectionString)
        m_Connexion = Connexion

    End Function

    Public Function Command(ByVal pstrSql As String) As Odbc.OdbcCommand

        Command = New Odbc.OdbcCommand(pstrSql, m_Connexion)

        m_Command = Command

    End Function

    Public Function Command(ByVal pstrSql As String, ByRef oConnexion As Odbc.OdbcConnection) As Odbc.OdbcCommand

        Command = New Odbc.OdbcCommand(pstrSql, oConnexion)
        m_Command = Command

    End Function

    Public Function DataAdapter(ByRef oCommand As Odbc.OdbcCommand) As Odbc.OdbcDataAdapter

        DataAdapter = New Odbc.OdbcDataAdapter(oCommand)
        'Call SetCommandBuilders()
        m_DataAdapter = DataAdapter

    End Function

    Public Function DataAdapter(ByVal pstrSql As String) As Odbc.OdbcDataAdapter

        DataAdapter = New Odbc.OdbcDataAdapter(Command(pstrSql))
        'Call SetCommandBuilders()
        m_DataAdapter = DataAdapter

    End Function

    Public Function DataSet(ByRef oDataAdapter As Odbc.OdbcDataAdapter) As DataSet

        DataSet = New DataSet
        oDataAdapter.Fill(DataSet)
        m_DataSet = DataSet

    End Function

    Public Function DataSet(ByVal pstrSql As String) As DataSet

        DataSet = New DataSet
        DataAdapter(pstrSql).Fill(DataSet)
        m_DataSet = DataSet

    End Function

    Public Function DataTable(ByRef oDataset As DataSet, ByVal pstrIndex As String) As DataTable

        DataTable = oDataset.Tables(pstrIndex)
        m_DataTable = DataTable

    End Function

    Public Function DataTable(ByRef oDataset As DataSet, ByVal piIndex As Integer) As DataTable
        DataTable = oDataset.Tables(piIndex)
        m_DataTable = DataTable
    End Function

    Public Function DataTable(ByVal pstrSql As String) As DataTable
        DataTable = DataSet(pstrSql).Tables(0)
        m_DataTable = DataTable
    End Function

    Public Sub Execute(ByVal pstrSql As String)

        m_Command = Command(pstrSql)

        If m_Command.Connection.State <> ConnectionState.Open Then m_Command.Connection.Open()

        m_Command.ExecuteNonQuery()

        If m_Command.Connection.State <> ConnectionState.Closed Then m_Command.Connection.Close()

    End Sub



    Const CONNECTION_STRING As String = "Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\NORTHWND.MDF;Integrated Security=True;User Instance=True"

    ' This function will be used to execute R(CRUD) operation of parameterless commands
    Friend Shared Function ExecuteSelectCommand() As DataTable
        Dim table As DataTable = Nothing
        Return table
    End Function

    ' This function will be used to execute R(CRUD) operation of parameterized commands
    Friend Shared Function ExecuteParamerizedSelectCommand() As DataTable
        Dim table As New DataTable()
        Return table
    End Function

    ' This function will be used to execute CUD(CRUD) operation of parameterized commands
    Friend Shared Function ExecuteNonQuery() As Boolean
        Dim result As Integer = 0
        Return (result > 0)
    End Function




End Class
