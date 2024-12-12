Option Strict Off
Option Explicit On
Imports System.Security.Principal
Imports System.Runtime.InteropServices

Friend Class frmLogin
    Inherits System.Windows.Forms.Form

    Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As String, _
        ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As LogonType, _
        ByVal dwLogonProvider As LogonProvider, ByRef phToken As IntPtr) As Integer

    Public Enum LogonType
        Interactive = 2
        Network
        Batch
        Servic
    End Enum

    Public Enum LogonProvider
        Defaut
        WinNT35
    End Enum

    'Public Settings As cSettings

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////       FORM LOAD PROCEDURES            /////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load


        lblVersion.Text = Version()
        Call FormLoad()
    End Sub
    Private Sub FormLoad()

        Dim tmpConnState As Integer  'Added June 2, 2017 - T. Louzon
        Try

            'Set default values for user
            'MsgBox("t :")
            gNTUserID = GetNTUserID()

            'Replaced Aug 3, 2017 - T. Louzon
            'SetDSNName(g_Default_Conn_String)
            If Trim(txtDSNName.Text) = "" Then
                gDSNName = GetSetting("OEI", "frmLogin", "txtDSNName", "EXACT")
                txtDSNName.Text = gDSNName
            Else
                txtDSNName.Text = UCase(txtDSNName.Text.Trim)
                gDSNName = txtDSNName.Text
            End If

            ' Get user from database. 
            'g_User = New cUser(gNTUserID, g_Default_Conn_String)
            g_User = New cUser(gNTUserID, "DSN=" & gDSNName)

            ' If User uses another database, reloads the user from the database.
            If g_User.Conn_String.Trim().ToUpper() <> g_Default_Conn_String.Trim().ToUpper() Then
                g_User = New cUser(gNTUserID, g_User.Conn_String)
                '''''    SetDSNName(g_User.Conn_String)   'Removed Aug 3, 2017
            End If

            m_oSettings = New cSettings()
            m_oOrder = New cOrder()
            g_oOrdline = New cOrdLine()

            ''==============================
            'Added June 28, 2017
            txtDatabaseName.Text = gsDataBaseName
            txtServerName.Text = gsServerName

            oFrmOrder = New frmOrder

            cmdOpen.Enabled = True
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Private Sub frmLogin_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        SaveSetting("OEI", "frmLogin", "txtDSNName", txtDSNName.Text)
    End Sub

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////       COMMAND BUTTON  PROCEDURES               /////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub cmdOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpen.Click

        Try

            'gVersion = lblVersion.Text
            'gNTUserID = GetNTUserID()
            'Call CreateConnection()
            'g_User = New cUser(gNTUserID)
            'oFrmOrder = New frmOrder

            Me.Visible = False
            oFrmOrder.ShowDialog()

            'WHEN DONE WITH frmORDER, close the application

            '++ID 05.18.2022 below line commented  it give us error message : Collection was modified; enumeration operation may not execute
            'and OEI going in Background proccess and didn't let open another session
            'in frmOrder was added FormClosing to close the form
            'oFrmOrder.Dispose()

            Me.Visible = False
            Me.Close()

            'free all
            Call ExitApplication()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
        Try
            'Close this form
            Me.Close()
            'Close all forms and end application
            Call ExitApplication()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtDSNName_TextChanged(sender As Object, e As EventArgs) Handles txtDSNName.TextChanged
        txtDatabaseName.Text = ""
        txtServerName.Text = ""
        cmdOpen.Enabled = False

    End Sub

    Private Sub txtDSNName_Leave(sender As Object, e As EventArgs) Handles txtDSNName.Leave
        Call FormLoad()
    End Sub





    '====================================================================================================
    '====================================================================================================
    '====================================================================================================
    '====================================================================================================
    '============  THIS CODE IS NOT USED AT THIS TIME ===================================================
    '====================================================================================================
    '====================================================================================================
    '====================================================================================================
    'Private Sub frmLogin_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '    'SaveSetting("Traveler", "Login", "txtDatabaseName", txtDatabaseName.Text)
    '    'SaveSetting("Traveler", "Login", "txtOEIServerName", txtServerName.Text)
    'End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'BUTTON1 is NOT VISIBLE - SO NOT USED

        Dim tokenHandle As New IntPtr(0)
        Dim wid_admin As WindowsIdentity = Nothing
        Dim wic As WindowsImpersonationContext = Nothing
        Const LOGON32_PROVIDER_DEFAULT As Integer = 0
        'Const LOGON32_LOGON_INTERACTIVE As Integer = 2
        'Const LOGON32_LOGON_SERVICE = 5

        tokenHandle = IntPtr.Zero

        Try
            Dim returnValue As Boolean = LogonUser("stephane", "BANKERS", "happy",
                              LogonType.Interactive, LOGON32_PROVIDER_DEFAULT, tokenHandle)

            'Me.Label1.Text = "Copying file..."
            wid_admin = New WindowsIdentity(tokenHandle)
            wic = wid_admin.Impersonate()
            'System.IO.File.Copy("c:\untitled.bmp", "\\\\domain\\shares\\Test\\untitled.bmp", True)
            System.IO.File.Copy("c:\untitled.bmp", "\\\\cobalt\\macolaes\\marc\\untitled.bmp", True)
            'Me.Label1.Text = "Copy succeeded"

        Catch se As System.Exception
            'Me.Label1.Text = "Copy Failed"
            Dim ret As Integer = Marshal.GetLastWin32Error()
            'Me.Label1.Text = (String.Format("Error code: {0} {1}", ret, se.Message))
        Finally
            If wic IsNot Nothing Then
                wic.Undo()
            End If
        End Try
    End Sub


    'Private Sub txtDatabaseName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDatabaseName.TextChanged

    '    Try

    '        gDatabaseName = txtDatabaseName.Text

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    '==============================
    'REMOVED June 2, 2017 - T. Louzon
    '==============================
    '''''Private Sub txtServerName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtServerName.TextChanged

    '''''    Try

    '''''        gServerName = txtServerName.Text

    '''''    Catch er As Exception
    '''''        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
    '''''    End Try

    '''''End Sub

    'Private Sub FormLoad()
    '    Dim tmpConnState As Integer  'Added June 2, 2017 - T. Louzon
    '    Try

    '        'Set default values for user
    '        'MsgBox("t :")
    '        gNTUserID = GetNTUserID()
    '        SetDSNName(g_Default_Conn_String)

    '        'Added June 1, 2017 - T. Louzon
    '        SetDSNName_Synergy(g_Default_ConnSynergy_String)

    '        ' Get user from database. 
    '        g_User = New cUser(gNTUserID, g_Default_Conn_String)

    '        ' If User uses another database, reloads the user from the database.
    '        If g_User.Conn_String.Trim().ToUpper() <> g_Default_Conn_String.Trim().ToUpper() Then
    '            g_User = New cUser(gNTUserID, g_User.Conn_String)
    '            SetDSNName(g_User.Conn_String)
    '        End If

    '        m_oSettings = New cSettings()
    '        m_oOrder = New cOrder()
    '        g_oOrdline = New cOrdLine()

    '        '==============================
    '        'Changed June 1, 2017 - T. Louzon
    '        '==============================
    '        txtDSNName.Text = m_oSettings.Login.DSNName ' GetSetting("Traveler", "Login", "txtServerName", "MacolaDB")
    '        txtDSNName_Synergy.Text = gDSNName_Synergy

    '        ''==============================
    '        ''Added June 2, 2017 - T. Louzon
    '        'tmpConnState = m_connexion.state
    '        'If m_connexion.state <> 1 Then
    '        '    m_connexion.state.open
    '        'End If
    '        'txtServerName = m_connexion.DataSource
    '        'txtDatabaseName = m_connexion.database
    '        'm_connexion.state = tmpConnState
    '        ''==============================
    '        'Added June 28, 2017
    '        txtDatabaseName.Text = gsDataBaseName
    '        txtServerName.Text = gsServerName

    '        oFrmOrder = New frmOrder

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(1).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

End Class

