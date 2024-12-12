Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices

Module mGlobal

    Public Declare Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As LogonType, ByVal dwLogonProvider As LogonProvider, ByRef phToken As IntPtr) As Integer
    Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Public Declare Function GetUsername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Public Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal enabled As Boolean) As Boolean
    Public Declare Function BringWindowToTop Lib "user32.dll" Alias "BringWindowToTop" (ByVal hwnd As Long) As Long
    Public Declare Function SetForegroundWindow Lib "user32.dll" Alias "SetForegroundWindow" (ByVal hwnd As IntPtr) As Boolean

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

    Public g_Default_Conn_String As String = "DSN=Exact;"
    Public g_Default_ConnSynergy_String As String = "DSN=Exact_ESE"

    Public oFrmOrder As frmOrder

    Public g_User As cUser
    Public m_oOrder As cOrder
    Public m_oSettings As cSettings

    Public g_oOrdline As cOrdLine
    Public g_oKitCollection As New Collection

    Public m_oOEMessage As New OEExceptionMessage
    Public g_Comp_Pref As String = "╘═ "

    '++ID 03.03.2020
    Public mglobal_changed_shipvia As Boolean

    Public g_ETA_Shortcut As String = "F3=ETA Details  "
    Public g_PriceBreaks_Shortcut As String = "F4=Price Breaks  "
    Public g_ItemInfo_Shortcut As String = "F5=Item Info  "
    Public g_ItemLocQty_Shortcut As String = "F6=Item/Loc Qty  "

    Public g_Search_Shortcut As String = "F7=Search  "
    Public g_Notes_Shortcut As String = "F8=Notes  "
    Public g_Date_Shortcut As String = "F9=Calendar  "
    Public g_Ship_Links_Shortcut As String = "F9=Order links  "
    Public g_Imprint_Shortcut As String = "Ctl-N=No Imprint  Ctl-Space=Empty  "
    Public g_Imprint_Location_Shortcut As String = "Ctl-B=Barrel  Ctl-C=Cap  Ctl-A=Enlarged Area  Ctl-S=Standard  Ctl-E=Enlarged  Ctl-F=Front  Ctl-Space=Empty  "
    Public g_Imprint_Color_Shortcut As String = "Ctl-B=Black  Ctl-S=Silver  Ctl-L=Laser  Ctl-W=White  Ctl-D=Deboss  Ctl-X=Laser/Oxid  Ctl-Space=Empty  "
    Public g_Imprint_Packaging_Shortcut As String = "Ctl-B=Bulk  Ctl-C=Cello  Ctl-S=Std  Ctl-P=P100  Ctl-W=Shrink Wrap  Ctl-Space=Empty  "
    Public g_Item_No_Shortcut As String = "Ctl-D=Deboss  Ctl-F=Paperproof  Ctl-S=Silk  Ctl-Q=Exact.Qty  Ctl-L=Laser.Setup  Ctl-C=Color  Ctl-R=Laser.OPM  Ctl-M=Silk.Metal  Ctl-Space=Empty  "
    Public g_Route_Shortcut As String = "Ctl-Space=Empty  "
    Public g_Order_Type_ShortCut As String = "B=Blanket  C=Credit Memo I=Invoice M=Master O=Order Q=Quote  "

    Public gDSNName As String
    '==============================
    'Added June 1, 2017 - T. Louzon - for use in frmLogin
    Public gDSNName_Synergy As String
    '==============================

    Public gNTUserID As String
    Public gGroupName As String

    Public RouteCollection As New Collection
    Public RouteIDcollection As New Collection

    Public gOrderNo As String
    Public gSelectedType As Short

    '==============================
    'Added June 1, 2017 - T. Louzon - for use in frmLogin
    Public gbchkLimitRoutes As Boolean    'Checkbox on frmOrder Tab page 3
    '==============================


    '==========================
    Public oTraveler_Reason_LateShipment As New cTraveler_Reason_LateShipment

    '--------------------------


    ' EXIT PROCEDURE
    Public Sub ExitApplication()

        Try
            '//////////////////////////////
            'Unload all of the forms
            Dim frm As System.Windows.Forms.Form
            For Each frm In My.Application.OpenForms
                frm.Close()
            Next frm

            'For ix As Integer = Application.OpenForms.Count - 1 To 0 Step -1
            '    Dim frm1 = Application.OpenForms(ix)
            '    frm1.Close()
            '    '' etc..
            'Next

            '///////////////////////////////////////
            'Delete the connection
            'UPGRADE_NOTE: Object gConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            'gConn = Nothing

            Application.Exit()
            End

            '////////
            'ErrorHandle:
            '            Select Case Err.Number
            '                Case 70 'attempting to delete files when one is open
            '                    Resume Next
            '                Case 53 'file not found when attempting delete (kill) - do nothing
            '                    Resume Next
            '                Case Else
            '                    MsgBox("Error Number:  " & Err.Number & vbCrLf & Err.Description, MsgBoxStyle.Critical, "Error in ' modGlobal_CT - ExitApplication' procedure ")
            '            End Select


        Catch er As Exception
            MsgBox("Error in mGlobal." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////       GET NTUSERID         //////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '++ID 4.25.2018
    Public Function Version() As String
        Version = ""

        Try

            Dim _assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly

            Dim fileInfo As New System.IO.FileInfo(_assembly.Location)

            Dim lastModified = fileInfo.LastWriteTime


            Version = "Version :" & " " & _assembly.GetName.Version.ToString & "; Date :  " & fileInfo.LastWriteTime 'lastModified.ToString("dd-MM-yyyy hh:mm tt") & ";"


        Catch ex As Exception
            MsgBox("Error in mGlobal." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function


    Public Function GetNTUserID() As String

        GetNTUserID = ""

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
            Dim str_Renamed As String
            Dim API_Return As Short
            Dim length As Integer

            'Create string of nulls
            str_Renamed = New String(Chr(0), 255)

            'Call API
            API_Return = GetUsername(str_Renamed, Len(str_Renamed))

            'Returns 0 if failure
            If API_Return = 0 Then
                'failure
                'UPGRADE_WARNING: Couldn't resolve default property of object GetNTUserID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GetNTUserID = ""
                Exit Function
            End If

            'str will have NTUserID followed by string of nulls
            'Find the first null characters in the string to determine the length of the userid
            length = InStr(1, str_Renamed, Chr(0)) - 1

            str_Renamed = Left(str_Renamed, length)

            'RETURN
            GetNTUserID = Trim(str_Renamed)
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

        Catch er As Exception
            MsgBox("Error in mGlobal." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    Public Function NoDate() As Date

        NoDate = New Date

    End Function



    Public Function SqlCompliantString(ByRef s As String) As String

        SqlCompliantString = Replace(s, "'", "''")

    End Function

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Sub CreateRouteCollections() ' , ByVal pstrProd_Cat As String)

        RouteCollection = New Collection
        RouteIDcollection = New Collection

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'Route must match category from line item.
            Dim SQL As String

            ' SQL Query to sort routes by group
            SQL = _
            "SELECT     DISTINCT RouteID, RouteDescription " & _
            "FROM       exact_Traveler_Route Rte where active = 1 " & _
            "ORDER BY   RouteDescription "

            Dim dt As New DataTable
            Dim db As New cDBA
            dt = db.DataTable(SQL)

            If dt.Rows.Count <> 0 Then
                For Each oRow As DataRow In dt.Rows
                    RouteCollection.Add(oRow.Item("RouteDescription").ToString.Trim, CStr(Trim(oRow.Item("RouteID"))))
                    RouteIDcollection.Add(CInt(oRow.Item("RouteID")), oRow.Item("RouteDescription").ToString.Trim)
                Next oRow
            End If

        Catch er As Exception
            MsgBox("Error in mGlobal." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Public Function ContactMethodString(ByVal pintContactMethod As ContactMethod) As String

        ContactMethodString = ""
        Select Case pintContactMethod
            Case ContactMethod.None
                ContactMethodString = ""
            Case ContactMethod.Email
                ContactMethodString = "E"
            Case ContactMethod.Fax
                ContactMethodString = "F"
        End Select

    End Function

    Public Function ContactMethodInt(ByVal pstrContactMethod As String) As ContactMethod

        ContactMethodInt = ContactMethod.None

        Select Case pstrContactMethod
            Case "E"
                ContactMethodInt = ContactMethod.Email
            Case "F"
                ContactMethodInt = ContactMethod.Fax
            Case Else
        End Select

    End Function


    Public Sub SetDSNName(ByVal pstrConn_String As String)
        gDSNName = pstrConn_String.Replace("DSN=", "").Replace(";", "").ToUpper
    End Sub

    'Added June 1, 2017 - T. Louzon
    Public Sub SetDSNName_Synergy(ByVal pstrConn_String As String)
        gDSNName_Synergy = pstrConn_String.Replace("DSN=", "").Replace(";", "").ToUpper
    End Sub
#Region "    DGV Columns functions ############################### "

    Public Function DGVButtonColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewButtonColumn

        DGVButtonColumn = New DataGridViewButtonColumn

        DGVButtonColumn.HeaderText = pstrHeaderText
        DGVButtonColumn.DataPropertyName = pstrName
        DGVButtonColumn.Name = pstrName

        If plWidth <> 0 Then DGVButtonColumn.Width = plWidth

    End Function

    Public Function DGVCalendarColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As CalendarColumn

        DGVCalendarColumn = New CalendarColumn

        DGVCalendarColumn.HeaderText = pstrHeaderText
        DGVCalendarColumn.DataPropertyName = pstrName
        DGVCalendarColumn.Name = pstrName
        'DGVCalendarColumn.DefaultCellStyle = "mm/dd/yyyy"
        DGVCalendarColumn.DefaultCellStyle.Format = "mm/dd/yyyy"

        If plWidth <> 0 Then DGVCalendarColumn.Width = plWidth

    End Function

    Public Function DGVCheckBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewCheckBoxColumn

        DGVCheckBoxColumn = New DataGridViewCheckBoxColumn

        DGVCheckBoxColumn.HeaderText = pstrHeaderText
        DGVCheckBoxColumn.DataPropertyName = pstrName
        DGVCheckBoxColumn.Name = pstrName

        DGVCheckBoxColumn.FalseValue = 0 '"False"
        DGVCheckBoxColumn.TrueValue = 1 '"True"
        DGVCheckBoxColumn.IndeterminateValue = 0 '"False"

        If plWidth <> 0 Then DGVCheckBoxColumn.Width = plWidth

    End Function

    Public Function DGVComboBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewComboBoxColumn

        DGVComboBoxColumn = New DataGridViewComboBoxColumn

        DGVComboBoxColumn.HeaderText = pstrHeaderText
        DGVComboBoxColumn.DataPropertyName = pstrName
        DGVComboBoxColumn.Name = pstrName
        'DGVComboBoxColumn.DropDownWidth = 160

        If plWidth <> 0 Then DGVComboBoxColumn.Width = plWidth

    End Function

    Public Function DGVImageColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, ByVal plWidth As Long) As DataGridViewImageColumn

        DGVImageColumn = New DataGridViewImageColumn
        DGVImageColumn.Name = pstrName
        DGVImageColumn.DataPropertyName = pstrName
        DGVImageColumn.HeaderText = pstrHeaderText

    End Function

    Public Function DGVTextBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 0) As DataGridViewTextBoxColumn

        DGVTextBoxColumn = New DataGridViewTextBoxColumn

        DGVTextBoxColumn.HeaderText = pstrHeaderText
        DGVTextBoxColumn.DataPropertyName = pstrName
        DGVTextBoxColumn.Name = pstrName

        If plWidth <> 0 Then DGVTextBoxColumn.Width = plWidth

    End Function

    Public Function GetDDTestsDataTable(ByVal pstrTable As String, ByVal pstrField As String) As DataTable

        Dim dt As DataTable
        Dim db As New cDBA
        Dim strsql As String = _
        "SELECT     DatabaseChar, DisplayChar, ISNULL(Description, '') AS Description " & _
        "FROM       DDTests WITH (nolock) " & _
        "WHERE      TableName = 'cicmpy' AND FieldName = 'cmp_type' " & _
        "ORDER BY   TableName, FieldName, DatabaseChar "

        dt = db.DataTable(strsql)

        Return dt

    End Function

    Public Function GetElementDescription(ByVal pstrTable As String, ByVal pstrField As String, ByVal pstrValue As String, Optional ByVal pstrValue2 As String = "") As String

        GetElementDescription = ""

        Dim dt As DataTable
        Dim db As New cDBA
        Dim strsql As String = ""

        Select Case pstrTable.ToUpper

            Case "ADDRESSSTATES" ' States and provinces
                strsql = _
                "SELECT StateCode, ISNULL(Name, '') AS Description, CountryCode " & _
                "FROM   AddressStates WITH (Nolock) " & _
                "WHERE  CountryCode='" & pstrValue & "' AND StateCode = '" & pstrValue2 & "' "

            Case "ARSLMFIL_SQL"
                strsql = "SELECT * FROM ARSLMFIL_SQL WITH (Nolock) WHERE HumRes_ID = '" & pstrValue & "'"

            Case "CICMPY"

                Select Case pstrField.ToUpper

                    Case "CMP_CODE"
                        strsql = _
                        "SELECT Cmp_Code, ISNULL(cmp_name, '') AS Description " & _
                        "FROM   Cicmpy WITH (Nolock) " & _
                        "WHERE 	Cmp_Code = '" & pstrValue & "' "

                    Case "CMP_PARENT"
                        strsql = _
                        "SELECT Cmp_Code AS Description " & _
                        "FROM   Cicmpy WITH (Nolock) " & _
                        "WHERE  CMP_WWN = '" & pstrValue & "' "

                    Case ""
                        'Case ""
                        'Case ""

                End Select

            Case "CLASSIFICATIONS"
                strsql = _
                "SELECT     C.ClassificationID, ISNULL(C.Description, '') AS Description " & _
                "FROM       Classifications C WITH (Nolock) " & _
                "LEFT JOIN  DDTests D WITH (Nolock) ON D.TableName = 'cicmpy' AND D.FieldName = 'cmp_type' AND C.Type = D.DatabaseChar " & _
                "WHERE      C.Type IS NULL AND C.ClassificationID = '" & pstrValue & "' "

            Case "LAND" ' Country
                strsql = _
                "SELECT LandCode, ISNULL(OMS60_0, '') AS Description " & _
                "FROM   Land WITH (Nolock) " & _
                "WHERE  Active = 1 AND LandCode = '" & pstrValue & "' "

            Case "SYCDEFIL_SQL"
                strsql = _
                "SELECT Sy_Code, ISNULL(Code_Desc, '') AS Description " & _
                "FROM   SYCDEFIL_SQL WITH (Nolock) " & _
                "WHERE  Sy_Code = '" & pstrValue & "'"

            Case "SYTRMFIL_SQL"
                strsql = _
                "SELECT Term_Code, ISNULL(Description, '') AS Description " & _
                "FROM   SYTRMFIL_SQL WITH (Nolock) " & _
                "WHERE  Term_Code = '" & pstrValue & "'"

            Case "TAAL"
                strsql = _
                "SELECT TaalCode, ISNULL(OMS30_0, '') AS Description " & _
                "FROM   Taal WITH (Nolock) " & _
                "WHERE  TaalCode = '" & pstrValue & "' "

            Case "TAXDETL_SQL"
                strsql = _
                "SELECT Tax_Cd, ISNULL(Tax_Cd_Description, '') AS Description " & _
                "FROM   TAXDETL_SQL WITH (Nolock) " & _
                "WHERE  Tax_Cd = '" & pstrValue & "' "

            Case "TAXSCHED_SQL"
                strsql = _
                "SELECT Tax_Sched, ISNULL(Tax_Sched_Desc, '') AS Description " & _
                "FROM   TAXSCHED_SQL WITH (Nolock) " & _
                "WHERE  Tax_Sched = '" & pstrValue & "'"

        End Select

        If strsql = "" Then

            strsql = _
            "SELECT     DatabaseChar, DisplayChar, ISNULL(Description, '') AS Description " & _
            "FROM       DDTests WITH (nolock) " & _
            "WHERE      TableName = '" & pstrTable & "' AND FieldName = '" & pstrField & "' AND DatabaseChar = '" & pstrValue & "' " & _
            "ORDER BY   TableName, FieldName, DatabaseChar "

        End If

        dt = db.DataTable(strsql)

        If dt.Rows.Count <> 0 Then
            GetElementDescription = dt.Rows(0).Item("Description").ToString
        End If

        'Dim strSql As String = "SELECT * FROM SYCDEFIL_SQL WITH (Nolock) WHERE sy_code = '" & pstrShip_Via_Cd & "'"
        'Dim strSql As String = "SELECT * FROM SYTRMFIL_SQL WITH (Nolock) WHERE Term_Code = '" & pstrAr_Terms_Cd & "'"
        'Dim strSql As String = "SELECT * FROM TAXSCHED_SQL WITH (Nolock) WHERE Tax_Sched = '" & pstrTax_Sched & "'"
        'Dim strSql As String = _
        '"SELECT ISNULL(Tax_Cd_Description, '') AS Tax_Cd_Description " & _
        '"FROM TAXDETL_SQL WITH (Nolock) WHERE Tax_Cd"
        'Dim strSql As String = "SELECT * FROM ARSLMFIL_SQL WITH (Nolock) WHERE HumRes_ID = '" & piSlspsn_No & "'"

    End Function

#End Region

#Region "    CalendarColumn Class ##############################################"

    Public Class CalendarColumn
        Inherits DataGridViewColumn

        Public Sub New()
            MyBase.New(New CalendarCell())
        End Sub

        Public Overrides Property CellTemplate() As DataGridViewCell
            Get
                Return MyBase.CellTemplate
            End Get
            Set(ByVal value As DataGridViewCell)

                ' Ensure that the cell used for the template is a CalendarCell.
                If (value IsNot Nothing) AndAlso _
                    Not value.GetType().IsAssignableFrom(GetType(CalendarCell)) _
                    Then
                    Throw New InvalidCastException("Must be a CalendarCell")
                End If
                MyBase.CellTemplate = value

            End Set
        End Property

    End Class

    Public Class CalendarCell
        Inherits DataGridViewTextBoxCell

        Public Sub New()
            ' Use the short date format.
            Me.Style.Format = "d"

        End Sub

        Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, _
            ByVal initialFormattedValue As Object, _
            ByVal dataGridViewCellStyle As DataGridViewCellStyle)

            ' Set the value of the editing control to the current cell value.
            MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, _
                dataGridViewCellStyle)

            Dim ctl As CalendarEditingControl = _
                CType(DataGridView.EditingControl, CalendarEditingControl)

            ' Use the default row value when Value property is null.
            If (Me.Value Is Nothing) Or (Me.Value.Equals(DBNull.Value)) Then
                ctl.Value = CType(Me.DefaultNewRowValue, DateTime)
            Else
                ctl.Value = CType(Me.Value, DateTime)
            End If
        End Sub

        Public Overrides ReadOnly Property EditType() As Type
            Get
                ' Return the type of the editing control that CalendarCell uses.
                Return GetType(CalendarEditingControl)
            End Get
        End Property

        Public Overrides ReadOnly Property ValueType() As Type
            Get
                ' Return the type of the value that CalendarCell contains.
                Return GetType(DateTime)
            End Get
        End Property

        Public Overrides ReadOnly Property DefaultNewRowValue() As Object
            Get
                ' Use the current date and time as the default value.
                Return DateTime.Now
            End Get
        End Property

    End Class

    Class CalendarEditingControl
        Inherits DateTimePicker
        Implements IDataGridViewEditingControl

        Private dataGridViewControl As DataGridView
        Private valueIsChanged As Boolean = False
        Private rowIndexNum As Integer

        Public Sub New()
            Me.Format = DateTimePickerFormat.Short
        End Sub

        Public Property EditingControlFormattedValue() As Object _
            Implements IDataGridViewEditingControl.EditingControlFormattedValue

            Get
                Return Me.Value.ToShortDateString()
            End Get

            Set(ByVal value As Object)
                Try
                    ' This will throw an exception of the string is 
                    ' null, empty, or not in the format of a date.
                    Me.Value = DateTime.Parse(CStr(value))
                Catch
                    ' In the case of an exception, just use the default
                    ' value so we're not left with a null value.
                    Me.Value = DateTime.Now
                End Try
            End Set

        End Property

        Public Function GetEditingControlFormattedValue(ByVal context _
            As DataGridViewDataErrorContexts) As Object _
            Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

            Return Me.Value.ToShortDateString()

        End Function

        Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As  _
            DataGridViewCellStyle) _
            Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

            Me.Font = dataGridViewCellStyle.Font
            Me.CalendarForeColor = dataGridViewCellStyle.ForeColor
            Me.CalendarMonthBackground = dataGridViewCellStyle.BackColor

        End Sub

        Public Property EditingControlRowIndex() As Integer _
            Implements IDataGridViewEditingControl.EditingControlRowIndex

            Get
                Return rowIndexNum
            End Get
            Set(ByVal value As Integer)
                rowIndexNum = value
            End Set

        End Property

        Public Function EditingControlWantsInputKey(ByVal key As Keys, _
            ByVal dataGridViewWantsInputKey As Boolean) As Boolean _
            Implements IDataGridViewEditingControl.EditingControlWantsInputKey

            ' Let the DateTimePicker handle the keys listed.
            Select Case key And Keys.KeyCode
                Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, _
                    Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp

                    Return True

                Case Else
                    Return Not dataGridViewWantsInputKey
            End Select

        End Function

        Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) _
            Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

            ' No preparation needs to be done.

        End Sub

        Public ReadOnly Property RepositionEditingControlOnValueChange() _
            As Boolean Implements _
            IDataGridViewEditingControl.RepositionEditingControlOnValueChange

            Get
                Return False
            End Get

        End Property

        Public Property EditingControlDataGridView() As DataGridView _
            Implements IDataGridViewEditingControl.EditingControlDataGridView

            Get
                Return dataGridViewControl
            End Get
            Set(ByVal value As DataGridView)
                dataGridViewControl = value
            End Set

        End Property

        Public Property EditingControlValueChanged() As Boolean _
            Implements IDataGridViewEditingControl.EditingControlValueChanged

            Get
                Return valueIsChanged
            End Get
            Set(ByVal value As Boolean)
                valueIsChanged = value
            End Set

        End Property

        Public ReadOnly Property EditingControlCursor() As Cursor _
            Implements IDataGridViewEditingControl.EditingPanelCursor

            Get
                Return MyBase.Cursor
            End Get

        End Property

        Protected Overrides Sub OnValueChanged(ByVal eventargs As EventArgs)

            ' Notify the DataGridView that the contents of the cell have changed.
            valueIsChanged = True
            Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
            MyBase.OnValueChanged(eventargs)

        End Sub

    End Class

#End Region

    Public Enum SqlMode

        SelectNoLock
        SelectUpdLock
        Insert
        Update
        Delete

    End Enum

    Public Enum LoginResultEnum
        Login
        Quit
    End Enum

    Public Enum DeleteStatus
        NotDeleted
        Deleted
    End Enum


    '------------------------------------------------------
    Public Function CheckProd_Cat(ByVal _prod_cat As String, ByVal _item_no As String) As Boolean

        CheckProd_Cat = False ' if is false this mean doesn't containe ORA prod categorie in Item line

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            'Route must match category from line item.
            Dim SQL As String

            ' SQL Query to sort routes by group
            SQL = " select * from ( " &
        " Select kit_feat_fg, prod_cat As  prod_cat_cd,item_no,0 As seq_no , '' AS comp_item_no, prod_cat from imitmidx_sql where item_no =  '" & _item_no & "'" &
        " And ISNULL(kit_feat_fg,'') <> 'K'  " &
        " union " &
        " Select  i.kit_feat_fg,i.prod_cat as  prod_cat_cd,i.item_no,k.seq_no,k.comp_item_no,ic.prod_cat  from imkitfil_sql k inner join imitmidx_sql i on  " &
        " k.item_no = i.item_no inner join imitmidx_sql ic on k.comp_item_no = ic.item_no " &
        " where i.kit_feat_fg = 'K' and k.item_no = '" & _item_no & "' ) as t  where prod_cat = '" & _prod_cat & "' order by item_no ,  seq_no  "

            Dim dt As New DataTable
            Dim db As New cDBA
            dt = db.DataTable(SQL)

            If dt.Rows.Count <> 0 Then
                CheckProd_Cat = True 'this mean one of the items containe ora prod categories
            End If

        Catch er As Exception
            MsgBox("Error in mGlobal." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function
    '-----------------------------------------------------


    Private _25percent As Double = 25  'applied to 3ST and lower
    Private _50percent As Double = 50  'applied to 4ST and higher
    Private _100percent As Double = 100
    Public Function Ret_Discount(ByVal m_OrderType As String, ByVal m_byRegion As String, ByVal m_star As String, Optional ByVal m_ItemNo As String = "",
                    Optional m_Prod_Cat As String = "", Optional m_Calc_Price As Integer = 0, Optional m_Mfg_Loc As String = "") As Double
        Ret_Discount = 0
        Try
            Dim p_price_discount As Double = 0

            Dim m_spec_instr As Decimal = 0.0
            ' MsgBox("I am in Ret_Discount .")


            'm_OrderType = m_oOrder.Ordhead.User_Def_Fld_1
            'm_byRegion = ""
            'm_star = m_oOrder.Ordhead.Customer.ClassificationId.ToUpper

            'm_Prod_Cat = "" ' come from insertt line in the function
            'm_Calc_Price = "" 'come from insert line

            'm_Mfg_Loc = m_oOrder.Ordhead.Mfg_Loc

            '++ID 8.27.2018 CHECK FOR SPEC_INSTRUCTION, IF DISCOUNT EXIST REPLACE DISCOUNT DEFINED.
            'modified in entire function 
                m_spec_instr = Return_Discount_By_OrderType()

            'ID 02.03.2023 IF discount from OEI_ITEM_SPEC_INSTRUCTION is not 0.0  us it, if is Zero(0) go to check  in Select case ticket #31065
            If m_spec_instr = 0.0 Then

                '2.11.2019 removed from exception (if) 44code, need to add discount,too.
                Select Case Trim(m_OrderType)
                    Case "SS"
                        '++ID 8.15.2019 if OEI_Disco.... added for manage Deboss,Deboss Brandbatch/Brandshield 44 , need to be charged
                        If OEI_Discount_Item_Exception(m_ItemNo) <> True Then

                            Select Case m_star
                                Case "1ST", "2ST", "3ST" '++ID 05.24.2024 was removed 0ST from SS/SB ticket 39501

                                    If m_byRegion.Contains("USA") Then
                                        ' If Trim(m_ItemNo).Contains("44") Or Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0 Then
                                        If (Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0) Then
                                            p_price_discount = 0
                                        Else
                                            'p_price_discount = _25percent
                                            If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _25percent '++ID 05.24.2024 _50percent removed and added 25% ticket #39501 '++ID 04.19.2023 Commented _25percent ticket 32762 
                                        End If
                                    Else
                                        p_price_discount = _25percent '++ID 05.24.2024 _50percent removed and added 25% ticket #39501  '++ID 04.19.2023 Commented 0 ticket 32762 0 'canadian
                                    End If

                                Case "4ST", "5ST", "6ST" '++ID 05.24.2024 commented 7ST becasue it doesn't exist no more used , "7ST" 

                                    If m_byRegion.Contains("USA") Then
                                        ' If Trim(m_ItemNo).Contains("44") Or Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0 Then
                                        If (Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0) Then
                                            p_price_discount = 0
                                        Else
                                            'p_price_discount = _50percent
                                            If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _50percent
                                        End If
                                    Else
                                        p_price_discount = _50percent '++ID 04.19.2023 Commented 0 ticket 32762  0 'canadian
                                    End If
                            End Select
                        End If
                    Case "SP"

                        '++ID 12.14.2021 exception for by pass Ora prod cat, rreturn must be 0 (zero)
                        '  If m_Prod_Cat <> "ORA" Then

                        '++ID 07.12.2023 commented ORA exception by Rusela request , for reference ticket 34477 Self-promo discount on Ora items
                        'from Tanya sugestion

                        '   If CheckProd_Cat("ORA", m_ItemNo) = False Then


                        Select Case m_star
                                Case "1ST", "2ST", "3ST"
                                    If m_byRegion.Contains("USA") Then
                                        '++ID put in comments 4.9.2019
                                        '   If m_Prod_Cat <> "PRC" And m_Calc_Price > 5 Then
                                        'p_price_discount = _25percent
                                        If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _25percent
                                        ' End If
                                    Else
                                        '  If Trim(m_ItemNo).Contains("44") Or Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0 Then
                                        If Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0 Then
                                            p_price_discount = 0
                                        Else
                                            'p_price_discount = _25percent
                                            If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _25percent 'canadian
                                        End If
                                    End If
                            Case "4ST", "5ST", "6ST" '++ID 05.24.2024 commented 7ST becasue it doesn't exist no more used , "7ST" 
                                If m_byRegion.Contains("USA") Then
                                        ' If m_Prod_Cat <> "PRC" And m_Calc_Price > 5 Then
                                        'p_price_discount = _50percent
                                        If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _50percent
                                        '   End If
                                    Else
                                        '  If Trim(m_ItemNo).Contains("44") Or Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0 Then
                                        If (Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0) Then
                                            p_price_discount = 0
                                        Else
                                            '++ID 8.21.2018 exception for Staples CND discount not 50% but 25%
                                            'If Trim(m_oOrder.Ordhead.Customer.textfield3) <> "STAPLES CND" Then
                                            '    p_price_discount = _50percent  'canadian
                                            'Else
                                            '    p_price_discount = _25percent  'canadian
                                            'End If

                                            If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _50percent

                                        End If
                                    End If
                            End Select

                    '    End If 'ORA exception IF

                    Case "SB"
                        '++ID 8.15.2019 if OEI_Disco.... added for manage Deboss,Deboss Brandbatch/Brandshield 44 , need to be charged
                        If OEI_Discount_Item_Exception(m_ItemNo) <> True Then

                            Select Case m_star
                                Case "1ST", "2ST", "3ST" '++ID 05.24.2024 was removed 0ST from SS/SB ticket 39501
                                    If m_byRegion.Contains("USA") Then
                                        'p_price_discount = 0
                                        If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _25percent '++ID 05.24.2024 _50percent removed and added 25% ticket #39501  '++ID 04.19.2023 Commented 0 ticket 32762 
                                    Else
                                        'p_price_discount = _25percent
                                        If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _25percent '++ID 05.24.2024 _50percent removed and added 25% ticket #39501 '++ID 04.19.2023 Commented _25percent ticket 32762'canadian
                                    End If

                                Case "4ST", "5ST", "6ST" '++ID 05.24.2024 commented 7ST becasue it doesn't exist no more used , "7ST" 

                                    If m_byRegion.Contains("USA") Then
                                        'p_price_discount = 0
                                        If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _50percent '++ID 04.19.2023 Commented 0 ticket 32762 
                                    Else
                                        ' If Trim(m_ItemNo).Contains("44") Or Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0 Then
                                        If (Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0) Then
                                            p_price_discount = 0
                                        Else
                                            'p_price_discount = _25percent
                                            If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _50percent '++ID mod 2.27.2019 -_25percent  'canadian
                                        End If
                                    End If
                            End Select

                        End If

                        'Case "RS"
                        '    Select Case m_star
                        '        Case "1ST", "2ST", "3ST"
                        '            If m_byRegion.Contains("USA") Then
                        '                If Trim(m_Mfg_Loc) = "3" And (m_Prod_Cat = "TEK" Or m_Prod_Cat = "BAG") And m_Calc_Price > 20 Then
                        '                    'p_price_discount = _25percent
                        '                    If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _25percent
                        '                Else
                        '                    p_price_discount = 0
                        '                End If
                        '            Else
                        '                'canadian
                        '                If Trim(m_Mfg_Loc) = "3" And (m_Prod_Cat = "TEK" Or m_Prod_Cat = "BAG") And m_Calc_Price > 20 And m_star <> "3ST" Then
                        '                    ' If Trim(m_ItemNo).Contains("44") Or Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0 Then
                        '                    If (Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0) Then
                        '                        p_price_discount = 0
                        '                    Else
                        '                        'p_price_discount = _50percent
                        '                        If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _50percent 'canadian
                        '                    End If
                        '                Else
                        '                    p_price_discount = 0 'canadian
                        '                End If
                        '            End If

                        '        Case "4ST", "5ST", "6ST" '++ID 05.24.2024 commented 7ST becasue it doesn't exist no more used , "7ST" 
                        '            If m_byRegion.Contains("USA") Then
                        '                If Trim(m_Mfg_Loc) = "3" And (m_Prod_Cat = "TEK" Or m_Prod_Cat = "BAG") And m_Calc_Price > 20 Then
                        '                    ' p_price_discount = _50percent
                        '                    If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _50percent
                        '                Else
                        '                    p_price_discount = 0
                        '                End If
                        '            Else
                        '                If Trim(m_Mfg_Loc) = "3" And (m_Prod_Cat = "TEK" Or m_Prod_Cat = "BAG") And m_Calc_Price > 20 Then
                        '                    ' If Trim(m_ItemNo).Contains("44") Or Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0 Then
                        '                    If (Trim(m_ItemNo).Contains("88") Or Trim(m_ItemNo).Length = 0) Then
                        '                        p_price_discount = 0
                        '                    Else
                        '                        'p_price_discount = _50percent
                        '                        If m_spec_instr <> 0 Then p_price_discount = m_spec_instr Else p_price_discount = _50percent  'canadian
                        '                    End If
                        '                Else
                        '                    p_price_discount = 0
                        '                End If

                        '            End If
                        '    End Select

                    Case Else

                End Select
            Else
                p_price_discount = m_spec_instr

            End If

            Return p_price_discount

        Catch ex As Exception
            MsgBox("Error in mGlobal.Ret_Discount." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Public Function Return_Discount_By_OrderType() As Decimal
        Return_Discount_By_OrderType = 0.0
        Try
            Dim discount As Decimal = 0.0

            Dim dt As DataTable
            Dim db As New cDBA
            'Trim(m_oOrder.Ordhead.Customer.textfield3) 
            Dim strSql As String =
              " SELECT SPEC_INSTRUCTION FROM OEI_ITEM_SPEC_INSTRUCTION WHERE DATA_FIELD = 'DISCOUNT'   " _
            & " AND (CUS_NO = '" & Trim(m_oOrder.Ordhead.Customer.cmp_code) & "' OR  " _
            & " CUS_GROUP =  '" & Trim(m_oOrder.Ordhead.Customer.textfield3) & "' OR " _
            & " (CUS_NO = '" & Trim(m_oOrder.Ordhead.Customer.cmp_code) & "' AND " _
            & " CUS_GROUP =  '" & Trim(m_oOrder.Ordhead.Customer.textfield3) & "'))  " _
            & " AND SHOW_MESSAGEBOX = 0 AND ORDER_TYPE = '" & Trim(m_oOrder.Ordhead.User_Def_Fld_2) & "' AND " _
            & " Cmp_Fctry = '" & Trim(m_oOrder.Ordhead.Customer.cmp_fctry) & "' " _
            & "   order by CUS_NO,CUS_GROUP  desc  "


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                If IsNumeric(Trim(dt.Rows(0).Item("SPEC_INSTRUCTION").ToString)) Then
                    discount = Convert.ToDecimal(Trim(dt.Rows(0).Item("SPEC_INSTRUCTION").ToString))
                End If

            End If

            Return discount

        Catch ex As Exception
            MsgBox("Error in mGlobal.Return_Discount_By_OrderType." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function

    Public Function Return_Prod_Cat(ByVal m_item_no As String) As String
        Return_Prod_Cat = ""
        Try
            Dim prod_cat As String = ""
            Dim dt As DataTable
            Dim db As New cDBA

            Dim strSql As String =
            "select PROD_CAT from imitmidx_sql  WITH (Nolock) where item_no = '" & Trim(m_item_no) & "'"

            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then
                prod_cat = Trim(dt.Rows(0).Item("PROD_CAT").ToString)
            End If

            Return prod_cat

        Catch ex As Exception
            MsgBox("Error in mGlobal.Return_Prod_Cat." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function


    Public Function OEI_Discount_Item_Exception(ByVal m_item_no As String) As Boolean
        OEI_Discount_Item_Exception = False
        Try

            Dim dt As DataTable
            Dim db As New cDBA

            Dim strSql As String =
                  "select ID,ENUM_CAT,ENUM_VALUE  from MDB_CFG_ENUM where ENUM_CAT = 'OEI_DISCOUNT_EXCLUSION' and ENUM_VALUE = '" & Trim(m_item_no) & "'"
            '"select item_no from [dbo].[OEI_Discount_Item_Exception]  WITH (Nolock) where item_no = '" & Trim(m_item_no) & "'"

            dt = db.DataTable(strSql)

            'in the code will be used like if is true this means item found in view
            If dt.Rows.Count <> 0 Then
                Return True
            End If

        Catch ex As Exception
            MsgBox("Error in mGlobal.Return_Prod_Cat." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function


    Private Function chargeExc() As Boolean
        chargeExc = False
        Try
            Dim exc As Boolean = False

            Dim listExc As New ArrayList
            listExc.Add("44")

            Dim lst As String = ""
            For Each lst In listExc

            Next

            Return exc

        Catch ex As Exception

        End Try

    End Function

    'function for check if exist exception for send Order Ack
    'return false this means is not in oei_item_spec_instruction , ask for send order ack
    'return true doesn't ask and doesn't sent
    'all used properties are global
    Public Function Return_Order_Ack_Exception() As Boolean
        Return_Order_Ack_Exception = False
        Try
            Dim m_ord_ack As Boolean
            m_ord_ack = False

            Dim dt As DataTable
            Dim db As New cDBA

            Dim strSql As String =
              " SELECT SPEC_INSTRUCTION FROM OEI_ITEM_SPEC_INSTRUCTION WHERE DATA_FIELD = 'ORD_ACK'   " _
            & " AND (CUS_NO = '" & Trim(m_oOrder.Ordhead.Customer.cmp_code) & "' OR  " _
            & " CUS_GROUP =  '" & Trim(m_oOrder.Ordhead.Customer.textfield3) & "') " _
            & " AND SHOW_MESSAGEBOX = 0  order by CUS_NO,CUS_GROUP  desc  "


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                m_ord_ack = True

            End If

            Return m_ord_ack

        Catch ex As Exception
            MsgBox("Error in mGlobal.Return_Order_Ack_Exception." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & ex.Message)
        End Try
    End Function


    '----------------------------------------------------------------------------------------------
    '------------- Excfeption by shipvia cd and classification id ---------------------------------
    '----------------------------------------------------------------------------------------------
    'used frmOrder.sstOrder.click tab 3 , in ucOrder no more use  Element_LostFocus  line 1434 ElseIf txtElement.name = "txtShip_Via_Cd" And txtElement.text <> String.Empty And m_oOrder.Ordhead.Customer.ClassificationId <> String.Empty Then
    Public Function spec_ship_charge(ByVal m_ShipViaCd As String, ByVal m_classification_id As String, ByVal m_User_Def_Fld_2 As String) As String

        '..fld_2 is order type 
        spec_ship_charge = ""

        Try

            Dim db As New cODBC
            Dim dt As DataTable
            Dim sql As String

            sql = "Select SPEC_INSTRUCTION,shipviacd,levelstar from OEI_ITEM_SPEC_INSTRUCTION " _
                & " where  SHOW_MESSAGEBOX = 1 and " _
                & " shipviacd =  '" & Trim(m_ShipViaCd) & "' and levelstar = '" & Trim(m_classification_id) & "' and ORDER_TYPE = '" & m_User_Def_Fld_2 & "'"

            dt = db.DataTable(sql)

            If dt.Rows.Count <> 0 Then

                Return dt.Rows(0).Item("SPEC_INSTRUCTION").ToString

            End If

        Catch er As Exception
            MsgBox("Error in mGlobal.spec_ship_charge." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function



End Module
