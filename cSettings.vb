Imports Microsoft.Win32

Public Class cSettings

    Public AddDoc As cAddDoc
    Public CreateTraveler As cCreateTraveler
    'Public CurrentSettings As cCurrentSettings
    Public ExtraFields As cExtraFields
    Public Header As cHeader
    Public Login As cLogin
    Public Traveler As cTraveler
    Public Export As cExport
    Public colSearchCollection As New Collection
    Public colOrderCollection As New Collection

    Public Class cCreateTraveler
        Public OrderNumber As String
    End Class

    Public Class cCurrentSettings

        Public DatabaseName As String
        Public DSNName As String

    End Class

    Public Class cExport

        Public SavePath As String
        Public FileExtension As String

    End Class

    Public Class cExtraFields

        Public KeepRecordset As String
        Public OrderNumber As String
        Public RecordsetLocked As String
        Public LineNo As String
        Public ItemNo As String
        Public LocationOfExe As String
        Public PwdUserName As String

    End Class

    Public Class cAddDoc

        Public DocType As String

    End Class

    Public Class cHeader

        Public Criteria1Combo As String ' cbocriteria1_text
        Public Criteria1Value As String
        Public Criteria2Enabled As Boolean
        Public Criteria2Combo As String
        Public Criteria2Value As String
        Public Criteria2Visible As Boolean
        Public AutoRefresh As Boolean
        Public History As Boolean
        Public InProduction As Boolean '-- ajouter 
        Public InProductionAndHistory As Boolean
        Public ShowFromCalendar As Boolean
        Public ShowToCalendar As Boolean
        Public ExcludeCompletedStep As Boolean
        Public ExcludeStepEnabled As Boolean
        Public ExcludeStep As String
        Public Width As Integer
        Public SearchWithAnd As Boolean
        Public SearchWithOr As Boolean
        Public Date1 As Boolean
        Public Date1Enabled As Boolean
        Public Date2 As Boolean
        Public Date2Enabled As Boolean
        Public FromDate As String
        Public FromDateEnabled As Boolean
        Public LastDays As Integer
        Public LastDaysEnabled As Boolean
        Public LastDaysCaption As String
        Public LastDaysCaptionEnabled As Boolean
        Public ToDate As String
        Public ToDateEnabled As Boolean
        Public SortBy As String
        Public SortAsc As String
        Public SortBy2 As String
        Public SortAsc2 As String
        Public Route As String
        Public OrderStep As String

    End Class

    Public Class cLogin

        Public DatabaseName As String
        Public DSNName As String = gDSNName

    End Class


    Public Class cTraveler

        Public extProgOrderNo As String
        Public SelectedOrderNo As String
        Public SelectedTravelerNo As String

    End Class

    Public Sub New()

        AddDoc = New cAddDoc
        CreateTraveler = New cCreateTraveler
        'CurrentSettings = New cCurrentSettings
        ExtraFields = New cExtraFields
        Header = New cHeader
        Login = New cLogin
        Traveler = New cTraveler
        Export = New cExport

        Call Load()

    End Sub

    Private Sub Load()

        Dim rkKey As RegistryKey

        rkKey = Nothing

        Try

            'gNTUserID = GetNTUserID()
            'g_User = New cUser(gNTUserID)

            'Dim db As New cDBA(g_MainConnectionString)
            'Dim dt As DataTable
            'Dim strsql As String = _
            '"SELECT * FROM OEI_USERS WITH (Nolock) WHERE USER_ID = '" & g_User.Usr_ID & "' "

            'dt = db.DataTable(strsql)
            'If dt.Rows.Count <> 0 Then
            '    Login.ServerName = dt.Rows(0).Item("Conn_String").ToString().Trim()
            '    If Login.ServerName <> g_User.Conn_String Then
            '        g_User = New cUser(gNTUserID, Login.ServerName)
            '    End If
            'End If

            '[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Traveler\Login]
            'rkKey = Registry.CurrentUser.OpenSubKey("Software\VB and VBA Program Settings\Traveler\Login")
            'Login.DatabaseName = GetSetting("Traveler", "Login", "txtDatabaseName", "100")
            'Login.ServerName = GetSetting("Traveler", "Login", "txtServerName", "Exact")
            'Login.ServerName = GetSetting("Traveler", "Login", "txtOEIServerName", "Exact")
            'Login.DatabaseName = rkKey.GetValue("txtDatabaseName", "100")
            'Login.ServerName = rkKey.GetValue("txtServerName", "thor")

        Catch er As Exception
            'Finally
            '    If Not (rkKey Is Nothing) Then rkKey.Close()
        End Try

        'Try
        '    '[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Traveler\Login]
        '    rkKey = Registry.CurrentUser.OpenSubKey("Software\VB and VBA Program Settings\Traveler\Login")
        '    'Login.DatabaseName = rkKey.GetValue("txtDatabaseName", "100")
        '    'Login.ServerName = rkKey.GetValue("txtServerName", "thor")
        'Catch er As Exception
        'Finally
        '    If Not (rkKey Is Nothing) Then rkKey.Close()
        'End Try
        'rkKey = Nothing

        'Try
        '    rkKey = Nothing
        '    rkKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Spector\Order Entry Interface 1.0")
        '    Export.SavePath = rkKey.GetValue("SavePath", "C:\")
        '    Export.FileExtension = rkKey.GetValue("FileExtension", ".xls")
        'Catch er As Exception
        'Finally
        '    If Not (rkKey Is Nothing) Then rkKey.Close()
        'End Try

        Try
            rkKey = Nothing
            rkKey = Registry.CurrentUser.OpenSubKey("Software\Exact\Browsers\Exact Globe 2000")
            Dim strSubKeyNames() As String = rkKey.GetSubKeyNames
            Dim rkSubKey As RegistryKey = Nothing
            Dim strSearch As String = ""
            ' We only want FJMACEXE keys to have @VisibleCols elements.
            For Each strSubKey As String In strSubKeyNames
                If strSubKey.Contains("[FJMACEXE:]") Then
                    Try
                        rkSubKey = Nothing
                        rkSubKey = Registry.CurrentUser.OpenSubKey("Software\Exact\Browsers\Exact Globe 2000\" & strSubKey)
                        strSearch = rkSubKey.GetValue("@VisibleCols").ToString
                        strSubKey = Mid(strSubKey, 12)
                        If Not colSearchCollection.Contains(Trim(strSubKey).ToUpper) Then
                            colSearchCollection.Add(strSearch, Trim(strSubKey).ToUpper)
                        End If

                        strSearch = rkSubKey.GetValue("").ToString
                        If strSearch <> "" Then
                            If Not colOrderCollection.Contains(Trim(strSubKey).ToUpper) Then
                                colOrderCollection.Add(strSearch, Trim(strSubKey).ToUpper)
                            End If
                        End If

                    Catch ex As Exception
                    Finally
                        If Not (rkKey Is Nothing) Then rkKey.Close()
                    End Try
                End If
            Next strSubKey

        Catch er As Exception
        Finally
            If Not (rkKey Is Nothing) Then rkKey.Close()
        End Try

    End Sub

End Class
