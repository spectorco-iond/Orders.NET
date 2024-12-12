Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices

Module mGlobal2

    '    Public Declare Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As LogonType, ByVal dwLogonProvider As LogonProvider, ByRef phToken As IntPtr) As Integer
    '    Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    '    Public Declare Function GetUsername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
    '    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    '    Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    '    Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    '    Public Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal enabled As Boolean) As Boolean
    '    Public Declare Function BringWindowToTop Lib "user32.dll" Alias "BringWindowToTop" (ByVal hwnd As Long) As Long
    '    Public Declare Function SetForegroundWindow Lib "user32.dll" Alias "SetForegroundWindow" (ByVal hwnd As IntPtr) As Boolean

    '    Public Enum LogonType
    '        Interactive = 2
    '        Network
    '        Batch
    '        Servic
    '    End Enum

    '    Public Enum LogonProvider
    '        Defaut
    '        WinNT35
    '    End Enum

    '    Public oFrmOrder As frmOrder

    '    Public g_User As cUser
    '    Public m_oOrder As New cOrder
    '    Public m_oSettings = New cSettings

    '    Public g_oOrdline As New cOrdLine()
    '    Public g_oKitCollection As New Collection

    '    Public m_oOEMessage As New OEExceptionMessage
    '    Public g_Comp_Pref As String = "╘═ "

    '    Public g_ETA_Shortcut As String = "F3=ETA Details  "
    '    Public g_PriceBreaks_Shortcut As String = "F4=Price Breaks  "
    '    Public g_ItemInfo_Shortcut As String = "F5=Item Info  "
    '    Public g_ItemLocQty_Shortcut As String = "F6=Item/Loc Qty  "

    '    Public g_Search_Shortcut As String = "F7=Search  "
    '    Public g_Notes_Shortcut As String = "F8=Notes  "
    '    Public g_Date_Shortcut As String = "F9=Calendar  "
    '    Public g_Ship_Links_Shortcut As String = "F9=Order links  "
    '    Public g_Imprint_Shortcut As String = "Ctl-N=No Imprint  Ctl-Space=Empty  "
    '    Public g_Imprint_Location_Shortcut As String = "Ctl-B=Barrel  Ctl-C=Cap  Ctl-A=Enlarged Area  Ctl-S=Standard  Ctl-E=Enlarged  Ctl-F=Front  Ctl-Space=Empty  "
    '    Public g_Imprint_Color_Shortcut As String = "Ctl-B=Black  Ctl-S=Silver  Ctl-L=Laser  Ctl-W=White  Ctl-D=Deboss  Ctl-X=Laser/Oxid  Ctl-Space=Empty  "
    '    Public g_Imprint_Packaging_Shortcut As String = "Ctl-B=Bulk  Ctl-C=Cello  Ctl-S=Std  Ctl-P=P100  Ctl-W=Shrink Wrap  Ctl-Space=Empty  "
    '    Public g_Item_No_Shortcut As String = "Ctl-D=Deboss  Ctl-F=Paperproof  Ctl-S=Silk  Ctl-Q=Exact Qty  Ctl-L=Laser Setup Charge Ctl-C=Color  Ctl-R=Laser OPM  Ctl-Space=Empty  "
    '    Public g_Route_Shortcut As String = "Ctl-Space=Empty  "
    '    Public g_Order_Type_ShortCut As String = "B=Blanket  C=Credit Memo I=Invoice M=Master O=Order Q=Quote  "

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


    '    Public Function ContactMethodString(ByVal pintContactMethod As ContactMethod) As String

    '        ContactMethodString = ""
    '        Select Case pintContactMethod
    '            Case ContactMethod.None
    '                ContactMethodString = ""
    '            Case ContactMethod.Email
    '                ContactMethodString = "E"
    '            Case ContactMethod.Fax
    '                ContactMethodString = "F"
    '        End Select

    '    End Function

    '    Public Function ContactMethodInt(ByVal pstrContactMethod As String) As ContactMethod

    '        ContactMethodInt = ContactMethod.None

    '        Select Case pstrContactMethod
    '            Case "E"
    '                ContactMethodInt = ContactMethod.Email
    '            Case "F"
    '                ContactMethodInt = ContactMethod.Fax
    '            Case Else
    '        End Select

    '    End Function

    '    Public Function Crop(ByVal source As Bitmap, ByVal section As Rectangle) As Bitmap
    '        ' An empty bitmap which will hold the cropped image
    '        Dim bmp As New Bitmap(section.Width, section.Height)
    '        Dim g As Graphics = Graphics.FromImage(bmp)
    '        ' Draw the given area (section) of the source image
    '        ' at location 0,0 on the empty bitmap (bmp)
    '        g.DrawImage(source, 0, 0, section, GraphicsUnit.Pixel)
    '        Return bmp
    '    End Function

    '    Public Function Crop1(ByVal image As Image, ByVal selection As Rectangle) As Image
    '        Dim bmp As Bitmap = TryCast(image, Bitmap)

    '        ' Check if it is a bitmap:
    '        If bmp Is Nothing Then
    '            Throw New ArgumentException("No valid bitmap")
    '        End If

    '        ' Crop the image:
    '        Dim cropBmp As Bitmap = bmp.Clone(selection, Imaging.PixelFormat.DontCare)

    '        ' Release the resources:
    '        image.Dispose()

    '        Return cropBmp

    '    End Function

    '#Region "    DGV Columns functions ############################### "

    '    Public Function DGVButtonColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewButtonColumn

    '        DGVButtonColumn = New DataGridViewButtonColumn

    '        DGVButtonColumn.HeaderText = pstrHeaderText
    '        DGVButtonColumn.DataPropertyName = pstrName
    '        DGVButtonColumn.Name = pstrName

    '        If plWidth <> 0 Then DGVButtonColumn.Width = plWidth

    '    End Function

    '    Public Function DGVCalendarColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As CalendarColumn

    '        DGVCalendarColumn = New CalendarColumn

    '        DGVCalendarColumn.HeaderText = pstrHeaderText
    '        DGVCalendarColumn.DataPropertyName = pstrName
    '        DGVCalendarColumn.Name = pstrName
    '        'DGVCalendarColumn.DefaultCellStyle = "mm/dd/yyyy"
    '        DGVCalendarColumn.DefaultCellStyle.Format = "mm/dd/yyyy"

    '        If plWidth <> 0 Then DGVCalendarColumn.Width = plWidth

    '    End Function

    '    Public Function DGVCheckBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewCheckBoxColumn

    '        DGVCheckBoxColumn = New DataGridViewCheckBoxColumn

    '        DGVCheckBoxColumn.HeaderText = pstrHeaderText
    '        DGVCheckBoxColumn.DataPropertyName = pstrName
    '        DGVCheckBoxColumn.Name = pstrName

    '        DGVCheckBoxColumn.FalseValue = 0 '"False"
    '        DGVCheckBoxColumn.TrueValue = 1 '"True"
    '        DGVCheckBoxColumn.IndeterminateValue = 0 '"False"

    '        If plWidth <> 0 Then DGVCheckBoxColumn.Width = plWidth

    '    End Function

    '    Public Function DGVComboBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewComboBoxColumn

    '        DGVComboBoxColumn = New DataGridViewComboBoxColumn

    '        DGVComboBoxColumn.HeaderText = pstrHeaderText
    '        DGVComboBoxColumn.DataPropertyName = pstrName
    '        DGVComboBoxColumn.Name = pstrName
    '        'DGVComboBoxColumn.DropDownWidth = 160

    '        If plWidth <> 0 Then DGVComboBoxColumn.Width = plWidth

    '    End Function

    '    Public Function DGVImageColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, ByVal plWidth As Long) As DataGridViewImageColumn

    '        DGVImageColumn = New DataGridViewImageColumn
    '        DGVImageColumn.Name = pstrName
    '        DGVImageColumn.DataPropertyName = pstrName
    '        DGVImageColumn.HeaderText = pstrHeaderText

    '    End Function

    '    Public Function DGVTextBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewTextBoxColumn

    '        DGVTextBoxColumn = New DataGridViewTextBoxColumn

    '        DGVTextBoxColumn.HeaderText = pstrHeaderText
    '        DGVTextBoxColumn.DataPropertyName = pstrName
    '        DGVTextBoxColumn.Name = pstrName

    '        If plWidth <> 0 Then DGVTextBoxColumn.Width = plWidth

    '    End Function

    '    Public Function GetDDTestsDataTable(ByVal pstrTable As String, ByVal pstrField As String) As DataTable

    '        Dim dt As DataTable
    '        Dim db As New cDBA
    '        Dim strsql As String = _
    '        "SELECT     DatabaseChar, DisplayChar, ISNULL(Description, '') AS Description " & _
    '        "FROM       DDTests WITH (nolock) " & _
    '        "WHERE      TableName = 'cicmpy' AND FieldName = 'cmp_type' " & _
    '        "ORDER BY   TableName, FieldName, DatabaseChar "

    '        dt = db.DataTable(strsql)

    '        Return dt

    '    End Function

    '    Public Function GetElementDescription(ByVal pstrTable As String, ByVal pstrField As String, ByVal pstrValue As String, Optional ByVal pstrValue2 As String = "") As String

    '        GetElementDescription = ""

    '        Dim dt As DataTable
    '        Dim db As New cDBA
    '        Dim strsql As String = ""

    '        Select Case pstrTable.ToUpper

    '            Case "ADDRESSSTATES" ' States and provinces
    '                strsql = _
    '                "SELECT StateCode, ISNULL(Name, '') AS Description, CountryCode " & _
    '                "FROM   AddressStates WITH (Nolock) " & _
    '                "WHERE  CountryCode='" & pstrValue & "' AND StateCode = '" & pstrValue2 & "' "

    '            Case "ARSLMFIL_SQL"
    '                strsql = "SELECT * FROM ARSLMFIL_SQL WITH (Nolock) WHERE HumRes_ID = '" & pstrValue & "'"

    '            Case "CICMPY"

    '                Select Case pstrField.ToUpper

    '                    Case "CMP_CODE"
    '                        strsql = _
    '                        "SELECT Cmp_Code, ISNULL(cmp_name, '') AS Description " & _
    '                        "FROM   Cicmpy WITH (Nolock) " & _
    '                        "WHERE 	Cmp_Code = '" & pstrValue & "' "

    '                    Case "CMP_PARENT"
    '                        strsql = _
    '                        "SELECT Cmp_Code AS Description " & _
    '                        "FROM   Cicmpy WITH (Nolock) " & _
    '                        "WHERE  CMP_WWN = '" & pstrValue & "' "

    '                    Case ""
    '                        'Case ""
    '                        'Case ""

    '                End Select

    '            Case "CLASSIFICATIONS"
    '                strsql = _
    '                "SELECT     C.ClassificationID, ISNULL(C.Description, '') AS Description " & _
    '                "FROM       Classifications C WITH (Nolock) " & _
    '                "LEFT JOIN  DDTests D WITH (Nolock) ON D.TableName = 'cicmpy' AND D.FieldName = 'cmp_type' AND C.Type = D.DatabaseChar " & _
    '                "WHERE      C.Type IS NULL AND C.ClassificationID = '" & pstrValue & "' "

    '            Case "LAND" ' Country
    '                strsql = _
    '                "SELECT LandCode, ISNULL(OMS60_0, '') AS Description " & _
    '                "FROM   Land WITH (Nolock) " & _
    '                "WHERE  Active = 1 AND LandCode = '" & pstrValue & "' "

    '            Case "SYCDEFIL_SQL"
    '                strsql = _
    '                "SELECT Sy_Code, ISNULL(Code_Desc, '') AS Description " & _
    '                "FROM   SYCDEFIL_SQL WITH (Nolock) " & _
    '                "WHERE  Sy_Code = '" & pstrValue & "'"

    '            Case "SYTRMFIL_SQL"
    '                strsql = _
    '                "SELECT Term_Code, ISNULL(Description, '') AS Description " & _
    '                "FROM   SYTRMFIL_SQL WITH (Nolock) " & _
    '                "WHERE  Term_Code = '" & pstrValue & "'"

    '            Case "TAAL"
    '                strsql = _
    '                "SELECT TaalCode, ISNULL(OMS30_0, '') AS Description " & _
    '                "FROM   Taal WITH (Nolock) " & _
    '                "WHERE  TaalCode = '" & pstrValue & "' "

    '            Case "TAXDETL_SQL"
    '                strsql = _
    '                "SELECT Tax_Cd, ISNULL(Tax_Cd_Description, '') AS Description " & _
    '                "FROM   TAXDETL_SQL WITH (Nolock) " & _
    '                "WHERE  Tax_Cd = '" & pstrValue & "' "

    '            Case "TAXSCHED_SQL"
    '                strsql = _
    '                "SELECT Tax_Sched, ISNULL(Tax_Sched_Desc, '') AS Description " & _
    '                "FROM   TAXSCHED_SQL WITH (Nolock) " & _
    '                "WHERE  Tax_Sched = '" & pstrValue & "'"

    '        End Select

    '        If strsql = "" Then

    '            strsql = _
    '            "SELECT     DatabaseChar, DisplayChar, ISNULL(Description, '') AS Description " & _
    '            "FROM       DDTests WITH (nolock) " & _
    '            "WHERE      TableName = '" & pstrTable & "' AND FieldName = '" & pstrField & "' AND DatabaseChar = '" & pstrValue & "' " & _
    '            "ORDER BY   TableName, FieldName, DatabaseChar "

    '        End If

    '        dt = db.DataTable(strsql)

    '        If dt.Rows.Count <> 0 Then
    '            GetElementDescription = dt.Rows(0).Item("Description").ToString
    '        End If

    '        'Dim strSql As String = "SELECT * FROM SYCDEFIL_SQL WITH (Nolock) WHERE sy_code = '" & pstrShip_Via_Cd & "'"
    '        'Dim strSql As String = "SELECT * FROM SYTRMFIL_SQL WITH (Nolock) WHERE Term_Code = '" & pstrAr_Terms_Cd & "'"
    '        'Dim strSql As String = "SELECT * FROM TAXSCHED_SQL WITH (Nolock) WHERE Tax_Sched = '" & pstrTax_Sched & "'"
    '        'Dim strSql As String = _
    '        '"SELECT ISNULL(Tax_Cd_Description, '') AS Tax_Cd_Description " & _
    '        '"FROM TAXDETL_SQL WITH (Nolock) WHERE Tax_Cd"
    '        'Dim strSql As String = "SELECT * FROM ARSLMFIL_SQL WITH (Nolock) WHERE HumRes_ID = '" & piSlspsn_No & "'"

    '    End Function

    '#End Region

    '#Region "    CalendarColumn Class ##############################################"

    '    Public Class CalendarColumn
    '        Inherits DataGridViewColumn

    '        Public Sub New()
    '            MyBase.New(New CalendarCell())
    '        End Sub

    '        Public Overrides Property CellTemplate() As DataGridViewCell
    '            Get
    '                Return MyBase.CellTemplate
    '            End Get
    '            Set(ByVal value As DataGridViewCell)

    '                ' Ensure that the cell used for the template is a CalendarCell.
    '                If (value IsNot Nothing) AndAlso _
    '                    Not value.GetType().IsAssignableFrom(GetType(CalendarCell)) _
    '                    Then
    '                    Throw New InvalidCastException("Must be a CalendarCell")
    '                End If
    '                MyBase.CellTemplate = value

    '            End Set
    '        End Property

    '    End Class

    '    Public Class CalendarCell
    '        Inherits DataGridViewTextBoxCell

    '        Public Sub New()
    '            ' Use the short date format.
    '            Me.Style.Format = "d"

    '        End Sub

    '        Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, _
    '            ByVal initialFormattedValue As Object, _
    '            ByVal dataGridViewCellStyle As DataGridViewCellStyle)

    '            ' Set the value of the editing control to the current cell value.
    '            MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, _
    '                dataGridViewCellStyle)

    '            Dim ctl As CalendarEditingControl = _
    '                CType(DataGridView.EditingControl, CalendarEditingControl)

    '            ' Use the default row value when Value property is null.
    '            If (Me.Value Is Nothing) Or (Me.Value.Equals(DBNull.Value)) Then
    '                ctl.Value = CType(Me.DefaultNewRowValue, DateTime)
    '            Else
    '                ctl.Value = CType(Me.Value, DateTime)
    '            End If
    '        End Sub

    '        Public Overrides ReadOnly Property EditType() As Type
    '            Get
    '                ' Return the type of the editing control that CalendarCell uses.
    '                Return GetType(CalendarEditingControl)
    '            End Get
    '        End Property

    '        Public Overrides ReadOnly Property ValueType() As Type
    '            Get
    '                ' Return the type of the value that CalendarCell contains.
    '                Return GetType(DateTime)
    '            End Get
    '        End Property

    '        Public Overrides ReadOnly Property DefaultNewRowValue() As Object
    '            Get
    '                ' Use the current date and time as the default value.
    '                Return DateTime.Now
    '            End Get
    '        End Property

    '    End Class

    '    Class CalendarEditingControl
    '        Inherits DateTimePicker
    '        Implements IDataGridViewEditingControl

    '        Private dataGridViewControl As DataGridView
    '        Private valueIsChanged As Boolean = False
    '        Private rowIndexNum As Integer

    '        Public Sub New()
    '            Me.Format = DateTimePickerFormat.Short
    '        End Sub

    '        Public Property EditingControlFormattedValue() As Object _
    '            Implements IDataGridViewEditingControl.EditingControlFormattedValue

    '            Get
    '                Return Me.Value.ToShortDateString()
    '            End Get

    '            Set(ByVal value As Object)
    '                Try
    '                    ' This will throw an exception of the string is 
    '                    ' null, empty, or not in the format of a date.
    '                    Me.Value = DateTime.Parse(CStr(value))
    '                Catch
    '                    ' In the case of an exception, just use the default
    '                    ' value so we're not left with a null value.
    '                    Me.Value = DateTime.Now
    '                End Try
    '            End Set

    '        End Property

    '        Public Function GetEditingControlFormattedValue(ByVal context _
    '            As DataGridViewDataErrorContexts) As Object _
    '            Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

    '            Return Me.Value.ToShortDateString()

    '        End Function

    '        Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As  _
    '            DataGridViewCellStyle) _
    '            Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

    '            Me.Font = dataGridViewCellStyle.Font
    '            Me.CalendarForeColor = dataGridViewCellStyle.ForeColor
    '            Me.CalendarMonthBackground = dataGridViewCellStyle.BackColor

    '        End Sub

    '        Public Property EditingControlRowIndex() As Integer _
    '            Implements IDataGridViewEditingControl.EditingControlRowIndex

    '            Get
    '                Return rowIndexNum
    '            End Get
    '            Set(ByVal value As Integer)
    '                rowIndexNum = value
    '            End Set

    '        End Property

    '        Public Function EditingControlWantsInputKey(ByVal key As Keys, _
    '            ByVal dataGridViewWantsInputKey As Boolean) As Boolean _
    '            Implements IDataGridViewEditingControl.EditingControlWantsInputKey

    '            ' Let the DateTimePicker handle the keys listed.
    '            Select Case key And Keys.KeyCode
    '                Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, _
    '                    Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp

    '                    Return True

    '                Case Else
    '                    Return Not dataGridViewWantsInputKey
    '            End Select

    '        End Function

    '        Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) _
    '            Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

    '            ' No preparation needs to be done.

    '        End Sub

    '        Public ReadOnly Property RepositionEditingControlOnValueChange() _
    '            As Boolean Implements _
    '            IDataGridViewEditingControl.RepositionEditingControlOnValueChange

    '            Get
    '                Return False
    '            End Get

    '        End Property

    '        Public Property EditingControlDataGridView() As DataGridView _
    '            Implements IDataGridViewEditingControl.EditingControlDataGridView

    '            Get
    '                Return dataGridViewControl
    '            End Get
    '            Set(ByVal value As DataGridView)
    '                dataGridViewControl = value
    '            End Set

    '        End Property

    '        Public Property EditingControlValueChanged() As Boolean _
    '            Implements IDataGridViewEditingControl.EditingControlValueChanged

    '            Get
    '                Return valueIsChanged
    '            End Get
    '            Set(ByVal value As Boolean)
    '                valueIsChanged = value
    '            End Set

    '        End Property

    '        Public ReadOnly Property EditingControlCursor() As Cursor _
    '            Implements IDataGridViewEditingControl.EditingPanelCursor

    '            Get
    '                Return MyBase.Cursor
    '            End Get

    '        End Property

    '        Protected Overrides Sub OnValueChanged(ByVal eventargs As EventArgs)

    '            ' Notify the DataGridView that the contents of the cell have changed.
    '            valueIsChanged = True
    '            Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
    '            MyBase.OnValueChanged(eventargs)

    '        End Sub

    '    End Class

    '#End Region

    '    Public Enum SqlMode

    '        SelectNoLock
    '        SelectUpdLock
    '        Insert
    '        Update
    '        Delete

    '    End Enum

    '    Public Enum LoginResultEnum
    '        Login
    '        Quit
    '    End Enum

    '    Public Enum DeleteStatus
    '        NotDeleted
    '        Deleted
    '    End Enum

End Module
