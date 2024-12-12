Imports System.Data.Odbc

Public Class cODBC

    'Private a1 As Odbc.OdbcConnection
    Private m_Connexion As Odbc.OdbcConnection
    Private m_Command As Odbc.OdbcCommand
    Private m_CommandBuilder As Odbc.OdbcCommandBuilder
    Private m_DataAdapter As Odbc.OdbcDataAdapter
    Private m_DataSet As DataSet
    Private m_DataTable As DataTable

    Public Sub New()

        m_Connexion = New Odbc.OdbcConnection
        m_Command = New Odbc.OdbcCommand
        m_CommandBuilder = New Odbc.OdbcCommandBuilder
        m_DataAdapter = New Odbc.OdbcDataAdapter
        m_DataSet = New DataSet
        m_DataTable = New DataTable

    End Sub

    Public Function Connexion() As Odbc.OdbcConnection

        'Connexion = Connexion("Driver={SQL Server};Server=Exact;Trusted_Connection=Yes;Database=100")
        Connexion = Connexion("DSN=Exact;")
        'Connexion = Connexion("Provider=SQLOdbc;" & gConn.ConnectionString)
        m_Connexion = Connexion

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

        Command = New Odbc.OdbcCommand(pstrSql, Connexion())
        m_Command = Command

    End Function

    Public Function Command(ByVal pstrSql As String, ByRef Connexion As Odbc.OdbcConnection) As Odbc.OdbcCommand

        Command = New Odbc.OdbcCommand(pstrSql, Connexion)
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

        'Try
        m_Command = Command(pstrSql)
        'm_Command.Connection = New Odbc.OdbcConnection
        'm_Command.Connection = m_Connexion
        If m_Command.Connection.State <> ConnectionState.Open Then m_Command.Connection.Open()
        m_Command.ExecuteNonQuery()
        If m_Command.Connection.State <> ConnectionState.Closed Then m_Command.Connection.Close()

        'Catch er As Exception
        'MsgBox("Error in cDBA." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & vbCrLf & vbCrLf & pstrSql)
        'End Try

    End Sub

    'Public Sub Execute(ByVal pstrSql As String)

    '    Try
    '        m_Command = Command(pstrSql)
    '        'm_Command.Connection = New Odbc.OdbcConnection
    '        'm_Command.Connection = m_Connexion
    '        If m_Command.Connection.State <> ConnectionState.Open Then m_Command.Connection.Open()
    '        m_Command.ExecuteNonQuery()
    '        If m_Command.Connection.State <> ConnectionState.Closed Then m_Command.Connection.Close()

    '    Catch er As Exception
    '        MsgBox("Error in cDBA." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & vbCrLf & vbCrLf & pstrSql)
    '    End Try

    'End Sub

    'Private Sub SetCommandBuilders()

    '    Try
    '        Dim sqlCommand As New Odbc.OdbcCommandBuilder(m_DataAdapter)
    '        m_DataAdapter.SelectCommand = m_Command
    '        ' If there are INNER JOINS< automatic SQL generation doesnt work
    '        If InStr(m_Command.CommandText.ToString, " JOIN ") = 0 Then
    '            m_DataAdapter.UpdateCommand = sqlCommand.GetUpdateCommand
    '            m_DataAdapter.DeleteCommand = sqlCommand.GetDeleteCommand
    '            m_DataAdapter.InsertCommand = sqlCommand.GetInsertCommand
    '        End If

    '    Catch er As Exception
    '        MsgBox("Error in CDBA." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

End Class
