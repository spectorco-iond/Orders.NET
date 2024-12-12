'Option Strict Off
'Option Explicit On

'Imports System.Runtime.InteropServices

Module modGlobal_CT
    '    'V2003.02.10
    '    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    '////////////////       API DECLARATIONS             ////////////////////////////////////////////////////////////////////////
    '    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    '    Public Declare Function GetUsername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
    '    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    '    Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    '    Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    '    Public Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal enabled As Boolean) As Boolean
    '    Public Declare Function BringWindowToTop Lib "user32.dll" Alias "BringWindowToTop" (ByVal hwnd As Long) As Long
    '    Public Declare Function SetForegroundWindow Lib "user32.dll" Alias "SetForegroundWindow" (ByVal hwnd As IntPtr) As Boolean

    '    'IntPtr FindWindow(string lpClassName, string lpWindowName); 
    '    'GLOBAL VARIABLES
    '    'Public gDatabaseName As String
    '    Public gServerName As String
    '    'Public gVersion As String

    '    Public gConn As ADODB.Connection
    '    'Public gConnNT03 As ADODB.Connection
    '    Public gNTUserID As String
    '    Public gGroupName As String

    '    'Public rsLineItems As ADODB.Recordset

    '    Public RouteCollection As New Collection
    '    Public RouteIDcollection As New Collection

    '    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    '////////////////       FORM POSITIONS             ///////////////////////////////////////////////////////////////////////
    '    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    Public Structure FormPos
    '        Dim Top As Integer
    '        'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    '        Dim Left_Renamed As Integer
    '        Dim Modified As Boolean
    '    End Structure

    '    Public frmAddDocPos As FormPos
    '    Public frmCreateTravelerPos As FormPos
    '    Public frmNotesPos As FormPos

    '    Public gCustomerCode As String
    '    Public gOrderNo As String
    '    Public gSelectedType As Short


    '    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    ' CREATE CONNECTION
    '    Public Sub CreateConnection()
    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '        On Error GoTo ErrHandle
    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '        Dim cnConnection As String
    '        'cnConnection = "DRIVER={SQL Server};SERVER=" & gServerName & ";DATABASE=" & gDatabaseName & ";Trusted_Connection=YES"


    '        cnConnection = "DSN=Exact;"

    '        gConn = New ADODB.Connection
    '        gConn.ConnectionString = cnConnection

    '        '////////////////////////
    '        'Test connection
    '        gConn.Open()
    '        gConn.Close()
    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '        Exit Sub
    'ErrHandle:
    '        Select Case Err.Number
    '            Case -2147467259
    '                'Incorrect Sql Server Name - pass to calling procedure
    '                Err.Raise(Err.Number)
    '            Case 3705 ' Connection already open - on gConn.open
    '                Resume Next
    '            Case Else
    '                MsgBox("Error Number:  " & Err.Number & vbCrLf & Err.Description, MsgBoxStyle.Critical, "Error in ' modGlobal_CT - CreateConnection ' procedure ")
    '                Err.Raise(Err.Number)
    '        End Select
    '    End Sub




    '    '    Public Sub CreateConnectionNT03()
    '    '        On Error GoTo ErrHandle
    '    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '    '        Dim cnConnection As String
    '    '        cnConnection = "DRIVER={SQL Server};SERVER=BankersNT03;DATABASE=traveler_tables;Trusted_Connection=YES"

    '    '        gConnNT03 = New ADODB.Connection
    '    '        'Set connection Timeout  -- Modified May 31, 2004

    '    '        gConnNT03.ConnectionString = cnConnection
    '    '        gConnNT03.ConnectionTimeout = 60
    '    '        gConnNT03.CommandTimeout = 60 ' Added Feb 5, 2008


    '    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '    '        'Test connection
    '    '        If gConnNT03.State <> ADODB.ObjectStateEnum.adStateOpen Then gConnNT03.Open()

    '    '        If gConnNT03.State <> ADODB.ObjectStateEnum.adStateClosed Then gConnNT03.Close()

    '    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '    '        Exit Sub
    '    'ErrHandle:
    '    '        Select Case Err.Number
    '    '            Case -2147467259
    '    '                'Incorrect Sql Server Name - pass to calling procedure
    '    '                Err.Raise(Err.Number)
    '    '            Case Else
    '    '                MsgBox("Error Number:  " & Err.Number & vbCrLf & Err.Description, MsgBoxStyle.Critical, "Error in ' modGlobal - CreateConnectionNT03 ' procedure ")
    '    '        End Select

    '    '    End Sub

    '    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    '    ' EXIT PROCEDURE
    '    Public Sub ExitApplication()
    '        On Error GoTo ErrorHandle

    '        '//////////////////////////////
    '        'Unload all of the forms
    '        Dim frm As System.Windows.Forms.Form
    '        For Each frm In My.Application.OpenForms
    '            frm.Close()
    '        Next frm

    '        '///////////////////////////////////////
    '        'Delete the connection
    '        'UPGRADE_NOTE: Object gConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '        gConn = Nothing
    '        End

    '        '////////
    'ErrorHandle:
    '        Select Case Err.Number
    '            Case 70 'attempting to delete files when one is open
    '                Resume Next
    '            Case 53 'file not found when attempting delete (kill) - do nothing
    '                Resume Next
    '            Case Else
    '                MsgBox("Error Number:  " & Err.Number & vbCrLf & Err.Description, MsgBoxStyle.Critical, "Error in ' modGlobal_CT - ExitApplication' procedure ")
    '        End Select
    '    End Sub


    '    '    Public Function isTravelerOnline() As Boolean
    '    '        On Error GoTo ErrHandle
    '    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '    '        Dim SQL As String
    '    '        Dim rsOnline As ADODB.Recordset

    '    '        Call CreateConnectionNT03()

    '    '        ' Is Traveler in Maintenance mode?
    '    '        SQL = " SELECT isInMaintenance " & " FROM traveler_config "
    '    '        rsOnline = New ADODB.Recordset
    '    '        rsOnline.Open(SQL, gConnNT03.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)

    '    '        If Not rsOnline.BOF And Not rsOnline.EOF Then
    '    '            If rsOnline.Fields("isInMaintenance").Value = 1 Then
    '    '                isTravelerOnline = False
    '    '            Else
    '    '                isTravelerOnline = True
    '    '            End If
    '    '        Else
    '    '            isTravelerOnline = False
    '    '        End If

    '    '        ' Clear Mem
    '    '        rsOnline.Close()
    '    '        'UPGRADE_NOTE: Object rsOnline may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '    '        rsOnline = Nothing
    '    '        'UPGRADE_NOTE: Object gConnNT03 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '    '        gConnNT03 = Nothing

    '    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '    '        Exit Function
    '    'ErrHandle:
    '    '        MsgBox("Error Number:  " & Err.Number & vbCrLf & Err.Description, MsgBoxStyle.Critical, "Error in ' modGlobal - SaveIPAddress ' procedure ")
    '    '    End Function



    '    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    '////////////////       GET NTUSERID         //////////////////////////////////////////////////////////////////////////////////
    '    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    Public Function GetNTUserID() As String

    '        'Returns NT UserID
    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '        On Error GoTo ErrHandle
    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '        'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    '        Dim str_Renamed As String
    '        Dim API_Return As Short
    '        Dim length As Integer

    '        'Create string of nulls
    '        str_Renamed = New String(Chr(0), 255)

    '        'Call API
    '        API_Return = GetUsername(str_Renamed, Len(str_Renamed))

    '        'Returns 0 if failure
    '        If API_Return = 0 Then
    '            'failure
    '            'UPGRADE_WARNING: Couldn't resolve default property of object GetNTUserID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '            GetNTUserID = ""
    '            Exit Function
    '        End If

    '        'str will have NTUserID followed by string of nulls
    '        'Find the first null characters in the string to determine the length of the userid
    '        length = InStr(1, str_Renamed, Chr(0)) - 1

    '        str_Renamed = Left(str_Renamed, length)

    '        'RETURN
    '        GetNTUserID = Trim(str_Renamed)
    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '        Exit Function
    'ErrHandle:
    '        MsgBox("Error Number:  " & Err.Number & vbCrLf & Err.Description, MsgBoxStyle.Critical, "Error in ' modUserName - GetNTUserID ' procedure ")
    '    End Function

    '    Public Function NoDate() As Date

    '        NoDate = New Date

    '    End Function

    '    'Public Function SearchElement(ByVal pstrElement As String, Optional ByVal pstrSearch As String = "") As Object

    '    '    Dim oForm As New frmSearch
    '    '    Try
    '    '        oForm.SearchElement = pstrElement
    '    '        oForm.SearchFromWindow = pstrSearch
    '    '        oForm.ShowDialog()
    '    '        SearchElement = oForm.FoundElementValue
    '    '    Catch er As Exception
    '    '        MsgBox(ex.Message)
    '    '        SearchElement = ""
    '    '    End Try

    '    '    'frmSearch.SearchElement = pstrElement
    '    '    'frmSearch.ShowDialog()

    '    'End Function

    '    Public Function SqlCompliantString(ByRef s As String) As String
    '        On Error GoTo ErrorHandle
    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '        SqlCompliantString = Replace(s, "'", "''")

    '        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '        Exit Function
    'ErrorHandle:
    '        MsgBox("Error Number:  " & Err.Number & vbCrLf & Err.Description, MsgBoxStyle.Critical, "Error in 'modGlobal - sqlFix'")
    '    End Function

    '    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '    Public Sub CreateRouteCollections() ' , ByVal pstrProd_Cat As String)

    '        RouteCollection = New Collection
    '        RouteIDcollection = New Collection

    '        Try
    '            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '            'Route must match category from line item.
    '            Dim SQL As String

    '            ' SQL Query to sort routes by group
    '            SQL = _
    '            "SELECT     DISTINCT RouteID, RouteDescription " & _
    '            "FROM       exact_Traveler_Route Rte where active = 1 " & _
    '            "ORDER BY   RouteDescription "

    '            Dim dt As New DataTable
    '            Dim db As New cDBA
    '            dt = db.DataTable(SQL)

    '            If dt.Rows.Count <> 0 Then
    '                For Each oRow As DataRow In dt.Rows
    '                    RouteCollection.Add(oRow.Item("RouteDescription").ToString, CStr(Trim(oRow.Item("RouteID"))))
    '                    RouteIDcollection.Add(CInt(oRow.Item("RouteID")), oRow.Item("RouteDescription").ToString)
    '                Next oRow
    '            End If

    '        Catch er As Exception
    '            MsgBox("Error in COrder." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '        End Try

    '    End Sub

End Module