Imports MySql.Data.MySqlClient

Public Class cMySqlDBA

    Private m_Connexion As MySqlConnection ' OleDb.OleDbConnection
    Private m_Command As MySqlCommand ' OleDb.OleDbCommand
    Private m_Reader As MySqlDataReader
    Private m_CommandBuilder As MySqlCommandBuilder '  OleDb.OleDbCommandBuilder
    Private m_DataAdapter As MySqlDataAdapter ' OleDb.OleDbDataAdapter
    Private m_DataSet As DataSet
    Private m_DataTable As DataTable

    'Dim con As MySqlConnection
    'Dim cmd As MySqlCommand
    'Dim reader As MySqlDataReader
    ''admbnkrs","bnk*0308
    'Dim _host As String = "www.spectorandco.ca" ' Connect to localhost database
    'Dim _database As String = "products2"
    'Dim _user As String = "admbnkrs" 'Enter your username, default is root
    'Dim _pass As String = "bnk*0308" 'Enter your password

    ''con = New MySqlConnection("Server=" + _host + ";User Id=" + _user + ";Password=" + _pass + ";")
    ''con = New MySqlConnection("SERVER=" + _host + ";DATABASE=" & _database & ";UID=" + _user + ";PASSWORD=" + _pass + ";")
    '    con = New MySqlConnection("SERVER=" + _host + ";UID=" + _user + ";PASSWORD=" + _pass + ";")
    '    Try
    '        con.Open()
    ''Check if the connection is open
    '        If con.State = ConnectionState.Open Then
    '            con.ChangeDatabase("products2") 'Use the MYSQL database for this example
    '            Console.WriteLine("Connection Open")
    'Dim Sql As String = "SELECT * FROM products_data_ca" ' Query the USER table to get user information
    '            cmd = New MySqlCommand(Sql, con)
    '            reader = cmd.ExecuteReader()

    ''Loop through all the users
    '            While reader.Read()
    '                Debug.Print("HOST: " & reader.GetString(0)) 'Get the host
    '                Debug.Print("USER: " & reader.GetString(1)) 'Get the username
    '                Debug.Print("PASS: " & reader.GetString(2)) 'Get the password
    '            End While
    '            reader.Close()

    '        End If


    '    Catch ex As Exception
    '        Debug.Print(ex.Message) ' Display any errors.
    '    End Try

    '    con.Close()


    Public Sub New()

        m_Connexion = New MySqlConnection
        m_Command = New MySqlCommand
        m_CommandBuilder = New MySqlCommandBuilder
        m_DataAdapter = New MySqlDataAdapter
        m_DataSet = New DataSet
        m_DataTable = New DataTable

    End Sub

    Public Function Connexion() As MySqlConnection

        Dim _host As String = "www.spectorandco.ca" ' Connect to localhost database
        Dim _database As String = "products2"
        Dim _user As String = "admbnkrs" 'Enter your username, default is root
        Dim _pass As String = "bnk*0308" 'Enter your password

        Connexion = New MySqlConnection("SERVER=" + _host + ";UID=" + _user + ";PASSWORD=" + _pass + ";")
        Try
            Connexion.Open()
            'Check if the connection is open
            If Connexion.State = ConnectionState.Open Then
                Connexion.ChangeDatabase("products2") 'Use the MYSQL database for this example
                '            Console.WriteLine("Connection Open")
                'Dim Sql As String = "SELECT * FROM products_data_ca" ' Query the USER table to get user information
                '            cmd = New MySqlCommand(Sql, con)
                '            reader = cmd.ExecuteReader()

                ''Loop through all the users
                '            While reader.Read()
                '                Debug.Print("HOST: " & reader.GetString(0)) 'Get the host
                '                Debug.Print("USER: " & reader.GetString(1)) 'Get the username
                '                Debug.Print("PASS: " & reader.GetString(2)) 'Get the password
                '            End While
                '            reader.Close()

            End If


        Catch ex As Exception
            Debug.Print(ex.Message) ' Display any errors.
        End Try

        'Connexion = Connexion("Provider=SQLOLEDB;" & gConn.ConnectionString)
        m_Connexion = Connexion

    End Function

    Public Property DBCommand() As MySqlCommand
        Get
            DBCommand = m_Command
        End Get
        Set(ByVal value As MySqlCommand)
            m_Command = value
        End Set
    End Property

    Public Property DBDataAdapter() As MySqlDataAdapter
        Get
            DBDataAdapter = m_DataAdapter
        End Get
        Set(ByVal value As MySqlDataAdapter)
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

    Public Function Connexion(ByVal pstrConnectionString) As MySqlConnection

        Connexion = New MySqlConnection(pstrConnectionString)
        m_Connexion = Connexion

    End Function

    Public Function Command(ByVal pstrSql As String) As MySqlCommand

        Command = New MySqlCommand(pstrSql, Connexion())
        m_Command = Command

    End Function

    Public Function Command(ByVal pstrSql As String, ByRef Connexion As MySqlConnection) As MySqlCommand

        Command = New MySqlCommand(pstrSql, Connexion)
        m_Command = Command

    End Function

    Public Function DataAdapter(ByRef oCommand As MySqlCommand) As MySqlDataAdapter

        DataAdapter = New MySqlDataAdapter(oCommand)
        'Call SetCommandBuilders()
        m_DataAdapter = DataAdapter

    End Function

    Public Function DataAdapter(ByVal pstrSql As String) As MySqlDataAdapter

        DataAdapter = New MySqlDataAdapter(Command(pstrSql))
        'Call SetCommandBuilders()
        m_DataAdapter = DataAdapter

    End Function

    Public Function DataSet(ByRef oDataAdapter As MySqlDataAdapter) As DataSet

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
        'm_Command.Connection = New OleDb.OleDbConnection
        'm_Command.Connection = m_Connexion
        If m_Command.Connection.State <> ConnectionState.Open Then m_Command.Connection.Open()
        m_Command.ExecuteNonQuery()
        If m_Command.Connection.State <> ConnectionState.Closed Then m_Command.Connection.Close()

        'Catch er As Exception
        'MsgBox("Error in cDBA." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & vbCrLf & vbCrLf & pstrSql)
        'End Try

    End Sub



End Class
