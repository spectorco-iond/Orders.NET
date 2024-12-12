Option Strict Off
Option Explicit On

Imports Microsoft.Win32
Imports System.Drawing

Friend Class frmSearch
    Inherits System.Windows.Forms.Form

    Public _FoundElementIndex As Object
    Public _FoundElementValue As Object
    Public _FoundRow As Integer

    Public _SearchElement As String
    Public _SearchFromWindow As String

    Private m_MinHeight As Integer
    Private m_MinWidth As Integer

    Private m_lLabels1Width As Integer
    Private m_lLabels2Width As Integer
    Private m_lText1Width As Integer
    Private m_lText2Width As Integer

    Private m_strSqlSearchQuery As Object
    Private m_strSqlSelect As String
    Private m_strSqlFields As String
    Private m_strSqlFrom As String
    Private m_SearchTable As String
    Private m_strSqlInnerJoin As String
    Private m_strSqlDefaultWhere As String
    Private m_strSqlSearchWhere As String
    Private m_strSqlOrderBy As String
    Private m_strSqlWhereElements As String
    Private m_UserName As String
    Private m_IndexField As String

    Private m_lElementCount As Integer
    Private m_lOrderByCount As Integer

    Private m_PrintPage As Integer
    Private m_lQtySearched As Integer

    Private m_FieldsCol As Collection
    Private m_SearchItemsCol As Collection
    'Private rsSearch As New ADODB.Recordset

    Private m_blnAll As Boolean = False

    ' Gets Key Down, to trigger ESC, F7, F8 and others
    Private Sub Pressed_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
        _txtElement_0.KeyDown, _txtElement_1.KeyDown, _txtElement_2.KeyDown, _
        _txtElement_3.KeyDown, _txtElement_4.KeyDown, _txtElement_5.KeyDown, _
        _txtElement_6.KeyDown, _txtElement_7.KeyDown, _txtElement_8.KeyDown, _
        _txtElement_9.KeyDown, _txtElement_10.KeyDown

        Dim txtElement As New Object

        Try

            If TypeOf sender Is xTextBox Then

                txtElement = New xTextBox
                txtElement = DirectCast(sender, xTextBox)

            End If

            Select Case e.KeyCode

                Case Keys.Return
                    cmdSearch.PerformClick()

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click

        Me.Close()

    End Sub

    Private Sub cmdNext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNext.Click

        Try
            If Not IsNumeric(txtTopRows.Text) Then Exit Sub

            If dgvSearch.Rows.Count < (CShort(txtTopRows.Text) - 1) Then Exit Sub

            Call GetSearch("Next", dgvSearch.Rows(dgvSearch.Rows.Count - 1).Cells(2).Value.ToString) ' fgSearch.get_TextMatrix(dgvSearch.rows.count- 1, 2))

            If dgvSearch.Rows.Count < 2 Then
                Call GetSearch("Previous", m_SearchItemsCol.Item(CStr(m_PrintPage)))
                If dgvSearch.Rows.Count < 2 Then Call cmdSearch_Click(cmdSearch, New System.EventArgs())
            Else
                m_PrintPage = m_PrintPage + 1
                m_SearchItemsCol.Add(dgvSearch.Rows(0).Cells(2).Value.ToString, CStr(m_PrintPage))
            End If

            dgvSearch.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdPrevious_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrevious.Click

        Try
            If m_PrintPage = 1 Then Exit Sub

            If dgvSearch.Rows.Count < 2 Then Exit Sub

            Call GetSearch("Previous", m_SearchItemsCol.Item(CStr(m_PrintPage - 1)))

            If dgvSearch.Rows.Count < 2 Then
                Call cmdSearch_Click(cmdSearch, New System.EventArgs())
            Else
                m_SearchItemsCol.Remove(CStr(m_PrintPage))
                m_PrintPage = m_PrintPage - 1
            End If

            dgvSearch.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdSearch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSearch.Click

        Dim lPos As Integer

        Try
            m_PrintPage = 1
            If m_SearchItemsCol.Count() <> 0 Then
                For lPos = m_SearchItemsCol.Count() To 1 Step -1
                    m_SearchItemsCol.Remove((lPos))
                Next lPos
            End If

            Call GetSearch()

            If dgvSearch.Rows.Count < 2 Then
                If dgvSearch.Rows.Count = 1 Then cmdSelect.Focus()
                Exit Sub
            End If
            m_SearchItemsCol.Add(dgvSearch.Rows(0).Cells(1).Value.ToString, CStr(m_PrintPage))

            If dgvSearch.Rows.Count > 1 Then
                dgvSearch.Focus()
            Else
                cmdSelect.Focus()
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub GetData(Optional ByVal pblnAll As Boolean = False)

        Dim iPos As Integer = 0
        Try
            m_blnAll = pblnAll

            m_MinHeight = 5320
            m_MinWidth = 6800
            m_lElementCount = 0
            m_lOrderByCount = 0

            m_lLabels1Width = 0
            m_lLabels2Width = 0
            m_lText1Width = 0
            m_lText2Width = 0

            m_SearchItemsCol = New Collection

            iPos = 1
            Call LoadSearch()

            iPos = 2
            Call ResizeElements()

            iPos = 3
            Call GetSearch()

            iPos = 4
            m_PrintPage = 1

            If dgvSearch.Rows.Count > 1 Then

                iPos = 5
                m_SearchItemsCol.Add(dgvSearch.Rows(0).Cells(1).Value.ToString, CStr(m_PrintPage))

            End If
            iPos = 6

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " Pos: " & iPos.ToString)
        End Try

    End Sub

    Private Sub Escape_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
                cmdClear.KeyDown, cmdClose.KeyDown, cmdNext.KeyDown, _
                cmdOpen.KeyDown, cmdPrevious.KeyDown, cmdSearch.KeyDown, _
                cmdSelect.KeyDown, txtElement.KeyDown, txtTopRows.KeyDown, dgvSearch.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose.PerformClick()
        End Select

    End Sub

    Private Sub frmSearch_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Try
            _FoundElementIndex = -1
            _FoundElementValue = String.Empty
            _FoundRow = -1

            Call GetData()

            _txtElement_1.Select()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'UPGRADE_WARNING: Event frmSearch.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub frmSearch_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize

        If VB6.PixelsToTwipsX(Me.Width) < m_MinWidth Then Me.Width = VB6.TwipsToPixelsX(m_MinWidth)
        If VB6.PixelsToTwipsY(Me.Height) < m_MinHeight Then Me.Height = VB6.TwipsToPixelsY(m_MinHeight)

        fraConfirmButtons.Left = (Me.Width - 300)
        fraConfirmButtons.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 1120)

        dgvSearch.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 360)
        dgvSearch.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(dgvSearch.Top) - 1105)

    End Sub

    Public Sub ResizeElements()

        Dim lElementSize As Integer
        Dim lPos As Integer
        Me.Refresh()

        Try
            m_MinHeight = 5320

            If lblElement(3).Tag <> "" Then m_MinHeight = 5740
            If lblElement(5).Tag <> "" Then m_MinHeight = 6160
            If lblElement(7).Tag <> "" Then m_MinHeight = 6580
            If lblElement(9).Tag <> "" Then m_MinHeight = 7000

            dgvSearch.Height = VB6.TwipsToPixelsY(2895)
            dgvSearch.Top = VB6.TwipsToPixelsY(m_MinHeight - 4000)

            For lPos = 1 To 10
                If _lblElement(lPos).Tag <> "" Then
                    ' set label width
                    'UPGRADE_ISSUE: Form method frmSearch.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    '                lElementSize = TextWidth(lblElement(lPos).Text)
                    lElementSize = TextWidth(lblElement(lPos).Text)
                    If lPos Mod 2 = 1 Then
                        If lElementSize > m_lLabels1Width Then m_lLabels1Width = lElementSize
                    Else
                        If lElementSize > m_lLabels2Width Then m_lLabels2Width = lElementSize
                    End If

                    ' set textbox width
                    'UPGRADE_WARNING: TextBox property txtElement.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                    lElementSize = VB6.TwipsToPixelsX(105 * txtElement(lPos).MaxLength) ' TextWidth(txtElement(lpos).Tag)
                    If lElementSize < VB6.TwipsToPixelsX((4 * 105)) Then lElementSize = VB6.TwipsToPixelsX((4 * 105)) ' SET MINIMUM ALWAYS 3 CHARACTERS
                    If lElementSize > VB6.TwipsToPixelsX((30 * 105)) Then lElementSize = VB6.TwipsToPixelsX((30 * 105)) ' SET MAXIMUM ALWAYS 30 CHARACTERS
                    If lPos Mod 2 = 1 Then
                        If (lElementSize + VB6.TwipsToPixelsX(120)) > m_lText1Width Then m_lText1Width = lElementSize + VB6.TwipsToPixelsX(120)
                    Else
                        If (lElementSize + VB6.TwipsToPixelsX(120)) > m_lText2Width Then m_lText2Width = lElementSize + VB6.TwipsToPixelsX(120)
                    End If
                    txtElement(lPos).Width = lElementSize + VB6.TwipsToPixelsX(120)
                End If
            Next lPos

            fraLeftElements.Left = fraLeftLabels.Left + m_lLabels1Width + VB6.TwipsToPixelsX(180)
            fraRightLabels.Left = fraLeftElements.Left + m_lText1Width + VB6.TwipsToPixelsX(180)
            fraRightElements.Left = fraRightLabels.Left + m_lLabels2Width + VB6.TwipsToPixelsX(180)

            Me.Width = fraRightElements.Left + m_lText2Width + VB6.TwipsToPixelsX(240)


            '========================================
            'Added May 31, 2017 - T. Louzon - Adding fields for Address to customer lookup 
            If _SearchElement = "OEI-CUSTOMER" Then
                Me.Width = 1000
            End If
            '========================================


            If Me.Width > m_MinWidth Then
                m_MinWidth = Me.Width
            End If

        Catch er As Exception
            MsgBox(er.Message)
        End Try

    End Sub

    Private Sub CreateSQLFieldsLine(ByRef pstrLine As String)

        Dim lPos As Long
        For lPos = 1 To m_lElementCount
            pstrLine = pstrLine & ", " & lblElement(lPos).Tag & " "
        Next lPos

    End Sub

    Private Sub CreateElements(ByVal pstrFields As String, ByVal pstrHeaders As String, ByVal pstrOrderBy As String, ByVal pstrSql As String)

        Dim rsFields As New ADODB.Recordset

        Dim strFields As String
        Dim strHeaders As String
        Dim strOrderBys As String

        strFields = pstrFields
        strHeaders = pstrHeaders
        strOrderBys = pstrOrderBy

        Dim lPosField As Integer
        Dim lPosHeader As Integer
        Dim lPosOrderBy As Integer

        Dim strField As String
        Dim strHeader As String
        Dim strOrderBy As String

        Try

            m_FieldsCol = New Collection

            While strOrderBys.Length > 0
                lPosOrderBy = InStr(1, strOrderBys, "|")
                If lPosOrderBy = 0 Then lPosOrderBy = InStr(1, strOrderBys, ",")

                If lPosOrderBy = 0 Then
                    strOrderBy = RTrim(strOrderBys)
                Else
                    strOrderBy = Mid(strOrderBys, 1, lPosOrderBy - 1)
                End If
                If Not (m_FieldsCol.Contains(Trim(strOrderBy.ToUpper))) Then
                    Call Add_SearchElement(strOrderBy, "", True) ' strOrderBy, strHeader, true
                    m_FieldsCol.Add(Trim(strOrderBy.ToUpper), Trim(strOrderBy.ToUpper))
                End If

                If lPosOrderBy = 0 Then
                    strOrderBys = ""
                Else
                    strOrderBys = Mid(strOrderBys, lPosOrderBy + 1)
                End If
            End While

            While Len(strFields) > 0
                lPosField = InStr(1, strFields, "|")
                If lPosField = 0 Then lPosField = InStr(1, strFields, ",")
                lPosHeader = InStr(1, strHeaders, "|")
                If lPosHeader = 0 Then lPosHeader = InStr(1, strHeaders, ",")
                If lPosField = 0 Then
                    strField = RTrim(strFields)
                    If strField.Contains(".") Then
                        strHeader = RTrim(strHeaders)
                    Else
                        strHeader = strField
                    End If

                Else
                    strField = Mid(strFields, 1, lPosField - 1)
                    If strField.Contains(".") Then
                        If lPosHeader = 0 Then
                            strHeader = Trim(strHeaders)
                        ElseIf lPosHeader = 1 Then
                            strHeader = ""
                        Else
                            strHeader = Mid(strHeaders, 1, lPosHeader - 1)
                        End If
                    Else
                        strHeader = strField
                    End If

                End If

                If Not (m_FieldsCol.Contains(Trim(strField.ToUpper))) Then
                    Call Add_SearchElement(strField, strHeader)
                    m_FieldsCol.Add(Trim(strField.ToUpper), Trim(strField.ToUpper))
                End If

                If lPosField = 0 Then
                    strFields = ""
                    strHeaders = ""
                Else
                    strFields = Mid(strFields, lPosField + 1)
                    If strField.Contains(".") Then
                        strHeaders = Mid(strHeaders, lPosHeader + 1)
                    End If
                End If

            End While

            'mb++ MAYBE NOT NECESSARY...
            'fgSearch.set_Cols(m_lElementCount + 1)
            Dim strSqlFields As String
            strSqlFields = "exec oei_efwGetDDInfo '" & m_SearchTable & "' "

            'rsFields.Open(pstrSql, Connection.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            rsFields.Open(strSqlFields, Connection.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            While Not rsFields.EOF
                'For Each rsField In rsFields.Fields
                For lPosField = 1 To m_lElementCount
                    If InStr(1, lblElement(lPosField).Tag.ToString.ToUpper & " ", "." & Trim(rsFields.Fields("ColumnName").Value.ToString.ToUpper) & " ") > 0 Then
                        'UPGRADE_WARNING: TextBox property txtElement.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                        txtElement(lPosField).MaxLength = CInt(rsFields.Fields("Length").Value)
                        txtElement(lPosField).Tag = rsFields.Fields("Type").Value
                        'If Trim(lblElement(lPosField).Text) = "" Then
                        lblElement(lPosField).Text = rsFields.Fields("Description").Value ' SetElementNameByFieldName(lblElement(lPosField).Tag)
                        'End If
                    ElseIf InStr(1, lblElement(lPosField).Tag.ToString.ToUpper & " ", " AS " & Trim(rsFields.Fields("ColumnName").Value.ToString.ToUpper) & " ") > 0 Then
                        'UPGRADE_WARNING: TextBox property txtElement.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                        txtElement(lPosField).MaxLength = CInt(rsFields.Fields("Length").Value)
                        txtElement(lPosField).Tag = rsFields.Fields("Type").Value
                        'If Trim(lblElement(lPosField).Text) = "" Then
                        lblElement(lPosField).Text = rsFields.Fields("Description").Value ' SetElementNameByFieldName(rsFields.Fields("ColumnName").ToString)
                        'Else
                        'lblElement(lPosField).Text = rsFields.Fields("Description").Value ' SetElementNameByFieldName(lblElement(lPosField).Text)
                        'End If
                    ElseIf InStr(1, lblElement(lPosField).Tag.ToString.ToUpper & " ", "." & Trim(rsFields.Fields("ColumnName").Value.ToString.ToUpper) & "_") > 0 Then
                        'UPGRADE_WARNING: TextBox property txtElement.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                        txtElement(lPosField).MaxLength = CInt(rsFields.Fields("Length").Value)
                        txtElement(lPosField).Tag = rsFields.Fields("Type").Value
                        'If Trim(lblElement(lPosField).Text) = "" Then
                        lblElement(lPosField).Text = rsFields.Fields("Description").Value ' SetElementNameByFieldName(rsFields.Fields("ColumnName").ToString)
                        'End If
                    ElseIf lblElement(lPosField).Tag.ToString.ToUpper = Trim(rsFields.Fields("ColumnName").Value.ToString.ToUpper) Then
                        'UPGRADE_WARNING: TextBox property txtElement.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                        txtElement(lPosField).MaxLength = CInt(rsFields.Fields("Length").Value)
                        txtElement(lPosField).Tag = rsFields.Fields("Type").Value
                        lblElement(lPosField).Text = rsFields.Fields("Description").Value ' SetElementNameByFieldName(rsFields.Fields("ColumnName").ToString)
                    End If

                Next lPosField
                'Next rsField
                rsFields.MoveNext()
            End While

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub Add_SearchElement(ByVal pstrField As String, ByVal pstrHeader As String, Optional ByVal pblnOrderBy As Boolean = False)

        Dim strFieldAS As String = ""

        Try
            If InStr(1, RTrim(UCase(pstrField)), " AS ") > 0 Then
                strFieldAS = Mid(pstrField, InStr(1, RTrim(UCase(pstrField)), " AS ") + 4)
            End If

            If RTrim(pstrHeader) = "" Then pstrHeader = strFieldAS

            m_lElementCount += 1

            ' We start by adding order by fields, they will appear first.
            If pblnOrderBy Then
                m_lOrderByCount += 1

                lblElement(m_lElementCount).Visible = True
                txtElement(m_lElementCount).Visible = True

                If InStr(1, RTrim(UCase(pstrField)), " AS ") > 0 Then
                    lblElement(m_lElementCount).Tag = Mid(pstrField, 1, InStr(1, RTrim(UCase(pstrField)), " AS ") - 1)
                Else
                    lblElement(m_lElementCount).Tag = pstrField
                End If

            Else

                lblElement(m_lElementCount).Visible = True
                txtElement(m_lElementCount).Visible = True

                pstrHeader = Replace(pstrHeader, "Ord ", "Order ")
                pstrHeader = Replace(pstrHeader, "Cus ", "Customer ")
                pstrHeader = Replace(pstrHeader, "Loc ", "Location ")

                lblElement(m_lElementCount).Text = RTrim(pstrHeader) & " "
                lblElement(m_lElementCount).Tag = pstrField

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadSearch()

        Dim rsSearch As ADODB.Recordset = New ADODB.Recordset
        Dim strSql As String

        Try
            strSql = "" & _
            "SELECT     Field_ID, Username, Srch_No, ISNULL(Pss_Title, '') AS Pss_Title, Srch_Table, " & _
            "           Columns1, Columns2, ISNULL(Headings1, '') AS Headings1, Headings2, " & _
            "           Column_Flags, ISNULL(Srch_Where, '') as Srch_Where, ISNULL(Srch_Joins, '') as Srch_Joins, " & _
            "           ISNULL(Srch_Order_By, '') as Srch_Order_By, ISNULL(Where_Defaults, '') AS Where_Defaults, " & _
            "           DB_Flag, Table_Pkg , Maint_Program, Maint_Prog_Type " & _
            "FROM       Searches_Sql WITH (Nolock) " & _
            "WHERE      Upper(RTrim(field_id)) = '" & _SearchElement & "' AND " & _
            "           (Upper(RTrim(username)) = 'MACOLA' )  AND srch_no = 1 " & _
            "UNION      " & _
            "SELECT     Field_ID, Username, Srch_No, ISNULL(Pss_Title, '') AS Pss_Title, Srch_Table, " & _
            "           Columns1, Columns2, ISNULL(Headings1, '') AS Headings1, Headings2, " & _
            "           Column_Flags, ISNULL(Srch_Where, '') as Srch_Where, ISNULL(Srch_Joins, '') as Srch_Joins, " & _
            "           ISNULL(Srch_Order_By, '') as Srch_Order_By, ISNULL(Where_Defaults, '') AS Where_Defaults, " & _
            "           DB_Flag, Table_Pkg , Maint_Program, Maint_Prog_Type " & _
            "FROM       OEI_SEARCHES_SQL WITH (Nolock) " & _
            "WHERE      Upper(RTrim(field_id)) = '" & _SearchElement & "' AND srch_no = 1 ORDER BY Srch_no "

            rsSearch.Open(strSql, Connection.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            '==================================================
            'Added May 30, 2017 - T. Louzon - To modify Customer search
            'IF not found add in search tables then add the new customer search
            If rsSearch.RecordCount = 0 And _SearchElement = "OEI-CUSTOMER" Then
                Dim sqlLOC As String
                sqlLOC = "INSERT INTO -(" & vbCrLf
                sqlLOC = sqlLOC & " field_id, username, srch_no, db_flag, pss_title, srch_table, table_pkg, columns1,  " & vbCrLf
                sqlLOC = sqlLOC & " columns2, headings1, headings2, column_flags, srch_where, srch_joins, srch_order_by,  " & vbCrLf
                sqlLOC = sqlLOC & " maint_program, maint_prog_type, where_defaults, filler_0001  " & vbCrLf
                sqlLOC = sqlLOC & " ) " & vbCrLf
                sqlLOC = sqlLOC & " SELECT " & vbCrLf
                sqlLOC = sqlLOC & " field_id = 'OEI-CUSTOMER', " & vbCrLf
                sqlLOC = sqlLOC & " username = 'MACOLA', " & vbCrLf
                sqlLOC = sqlLOC & " srch_no = '1', " & vbCrLf
                sqlLOC = sqlLOC & " db_flag = 'A', " & vbCrLf
                sqlLOC = sqlLOC & " pss_title = 'Customer Number', " & vbCrLf
                sqlLOC = sqlLOC & " srch_table = 'arcusfil_Sql', " & vbCrLf
                sqlLOC = sqlLOC & " table_pkg = 'AR', " & vbCrLf
                sqlLOC = sqlLOC & " columns1 = 'arcusfil_sql.cus_no|arcusfil_sql.cus_name|arcusfil_sql.hold_fg|arcusfil_Sql.city|arcusfil_sql.State|arcusfil_sql.zip   ', " & vbCrLf
                sqlLOC = sqlLOC & " columns2 = NULL, " & vbCrLf
                sqlLOC = sqlLOC & " headings1 = 'Cus No|Cus Name|Hold_fg|City|State|Zip', " & vbCrLf
                sqlLOC = sqlLOC & " headings2 = NULL, " & vbCrLf
                sqlLOC = sqlLOC & " column_flags = 'AR|A|A|A|A|A                                                                                        ', " & vbCrLf
                sqlLOC = sqlLOC & " srch_where = NULL, " & vbCrLf
                sqlLOC = sqlLOC & " srch_joins = NULL, " & vbCrLf
                sqlLOC = sqlLOC & " srch_order_by = 'arcusfil_Sql.cus_no', " & vbCrLf
                sqlLOC = sqlLOC & " maint_program = 'debitr', " & vbCrLf
                sqlLOC = sqlLOC & " maint_prog_type = '2', " & vbCrLf
                sqlLOC = sqlLOC & " where_defaults = NULL, " & vbCrLf
                sqlLOC = sqlLOC & " filler_0001 = NULL " & vbCrLf

                Dim db As New cDBA
                'ADD THE RECORD TO THE OEI_SEARCHES
                db.Execute(sqlLOC)
                'TRY TO OPEN AGAIN
                If rsSearch.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsSearch.Close()
                End If
                rsSearch.Open(strSql, Connection.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            End If
            '==================================================   End Changes - May 30, 2017 - T. Louzon

            '==================================================
            'Added July 17, 2017 - T. Louzon - To modify Item Search
            'IF not found add in search tables then add the new Item search
            'NOT MFG ITEM
            If rsSearch.RecordCount = 0 And _SearchElement = "OEI-IM-ITEM" Then
                Dim sqlLOC As String
                sqlLOC = "INSERT INTO OEI_SEARCHES_SQL(" & vbCrLf
                sqlLOC = sqlLOC & " field_id, username, srch_no, db_flag, pss_title, srch_table, table_pkg, columns1,  " & vbCrLf
                sqlLOC = sqlLOC & " columns2, headings1, headings2, column_flags, srch_where, srch_joins, srch_order_by,  " & vbCrLf
                sqlLOC = sqlLOC & " maint_program, maint_prog_type, where_defaults, filler_0001  " & vbCrLf
                sqlLOC = sqlLOC & " ) " & vbCrLf
                sqlLOC = sqlLOC & " SELECT " & vbCrLf
                sqlLOC = sqlLOC & " field_id = 'OEI-IM-ITEM', " & vbCrLf
                sqlLOC = sqlLOC & " username = 'MACOLA', " & vbCrLf
                sqlLOC = sqlLOC & " srch_no = '1', " & vbCrLf
                sqlLOC = sqlLOC & " db_flag = 'A', " & vbCrLf
                sqlLOC = sqlLOC & " pss_title = 'Item Number', " & vbCrLf
                sqlLOC = sqlLOC & " srch_table = 'IMITMIDX_SQL ', " & vbCrLf
                sqlLOC = sqlLOC & " table_pkg = 'IM', " & vbCrLf
                sqlLOC = sqlLOC & " columns1 = 'IMITMIDX_SQL.item_desc_1|IMITMIDX_SQL.item_no   ', " & vbCrLf
                sqlLOC = sqlLOC & " columns2 = NULL, " & vbCrLf
                sqlLOC = sqlLOC & " headings1 = 'Description|Item No ', " & vbCrLf
                sqlLOC = sqlLOC & " headings2 = NULL, " & vbCrLf
                sqlLOC = sqlLOC & " column_flags = 'A|AR ', " & vbCrLf
                sqlLOC = sqlLOC & " srch_where = 'pur_or_mfg <> ''M'''    , " & vbCrLf
                sqlLOC = sqlLOC & " srch_joins = NULL, " & vbCrLf
                sqlLOC = sqlLOC & " srch_order_by = 'IMITMIDX_SQL.item_no               ', " & vbCrLf
                sqlLOC = sqlLOC & " maint_program = 'IM0101', " & vbCrLf
                sqlLOC = sqlLOC & " maint_prog_type = '1', " & vbCrLf
                sqlLOC = sqlLOC & " where_defaults = NULL, " & vbCrLf
                sqlLOC = sqlLOC & " filler_0001 = NULL " & vbCrLf

                Dim db As New cDBA
                'ADD THE RECORD TO THE OEI_SEARCHES
                db.Execute(sqlLOC)
                'TRY TO OPEN AGAIN
                If rsSearch.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsSearch.Close()
                End If
                rsSearch.Open(strSql, Connection.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            End If
            '==================================================   End Changes - July 13, 2017 - T. Louzon



            If rsSearch.RecordCount <> 0 Then
                Dim strCustomSearch As String = ""
                Dim strOrderSearch As String = ""

                Me.Text = rsSearch.Fields("Pss_Title").Value

                m_SearchTable = Trim(rsSearch.Fields("Srch_Table").Value)
                m_strSqlFrom = " FROM " & RTrim(rsSearch.Fields("Srch_Table").Value) & " WITH (Nolock) "
                If Trim(rsSearch.Fields("UserName").Value) = "NO-ID" Then
                    'If Trim(rsSearch.Fields("Alternate_ID_Field").Value) <> "" Then ' Alternate_ID_Field
                    'm_strSqlFields = RTrim(rsSearch.Fields("Srch_Table").Value) & "." & Trim(rsSearch.Fields("Alternate_ID_Field").Value) & " AS ID, " & Replace(RTrim(rsSearch.Fields("Columns1").Value), "|", ",")
                    'Else
                    m_strSqlFields = RTrim(rsSearch.Fields("Srch_Order_By").Value) & " AS ID, " & Replace(RTrim(rsSearch.Fields("Columns1").Value), "|", ",")
                    'End If
                Else
                    m_strSqlFields = RTrim(rsSearch.Fields("Srch_Table").Value) & ".ID AS ID, " & Replace(RTrim(rsSearch.Fields("Columns1").Value), "|", ",")
                End If
                m_strSqlInnerJoin = RTrim(rsSearch.Fields("Srch_Joins").Value)
                m_strSqlDefaultWhere = RTrim(rsSearch.Fields("Where_Defaults").Value)

                Select Case _SearchElement
                    Case "SY-ACCOUNT2"
                        m_strSqlSearchWhere = "ksdrek.reknr = '     4000' "
                    Case "SY-ACCOUNT3"
                        m_strSqlSearchWhere = "ksprek.reknr = '     4000' "
                    Case Else
                        m_strSqlSearchWhere = RTrim(rsSearch.Fields("Srch_Where").Value)
                End Select

                Dim strCustomOrder As String = ""
                strCustomOrder = CustomOrder(RTrim(rsSearch.Fields("Srch_Table").Value))

                If strCustomOrder <> "" Then
                    m_strSqlOrderBy = " ORDER BY " & strCustomOrder
                Else
                    m_strSqlOrderBy = " ORDER BY " & RTrim(rsSearch.Fields("Srch_Order_By").Value)
                End If
                'm_strSqlOrderBy = " ORDER BY " & RTrim(rsSearch.Fields("Srch_Order_By").Value)

                m_UserName = rsSearch.Fields("UserName").Value

                If Trim(m_UserName) = "NO-ID" Then
                    strSql = "SELECT TOP 1 " & m_strSqlFields & m_strSqlFrom & m_strSqlInnerJoin
                Else
                    strSql = "SELECT TOP 1 " & m_strSqlFields & m_strSqlFrom & m_strSqlInnerJoin & " WHERE " & RTrim(rsSearch.Fields("Srch_Table").Value) & ".ID = 0"
                End If

                Call IndexField(rsSearch.Fields("Columns1").Value.ToString, rsSearch.Fields("Srch_Order_By").Value.ToString)

                If m_oSettings.colSearchCollection Is Nothing Then
                    Call CreateElements(rsSearch.Fields("Columns1").Value, rsSearch.Fields("Headings1").Value, rsSearch.Fields("Srch_Order_By").Value, strSql)
                ElseIf m_oSettings.colSearchCollection.Contains(Trim(rsSearch.Fields("Srch_Table").Value).ToUpper) And Mid(rsSearch.Fields("Field_ID").Value, 1, 3) <> "OEI" Then

                    If Trim(rsSearch.Fields("Field_ID").Value) = "SY-ACCOUNT2" Then

                        Call CreateElements(rsSearch.Fields("Columns1").Value, rsSearch.Fields("Headings1").Value, rsSearch.Fields("Srch_Order_By").Value, strSql)

                    Else

                        strCustomSearch = m_oSettings.colSearchCollection.Item(Trim(rsSearch.Fields("Srch_Table").Value).ToUpper).ToString.Replace(",", "|")
                        m_UserName = "NO-ID"

                        If m_oSettings.colOrderCollection Is Nothing Then
                            Call CreateElements(strCustomSearch, "", rsSearch.Fields("Srch_Order_By").Value, strSql)
                        ElseIf m_oSettings.colOrderCollection.Contains(Trim(rsSearch.Fields("Srch_Table").Value).ToUpper) Then
                            strOrderSearch = m_oSettings.colOrderCollection.Item(Trim(rsSearch.Fields("Srch_Table").Value).ToUpper).ToString
                            m_UserName = "ORDER"
                            Call CreateElements(strCustomSearch, "", strOrderSearch, strSql)
                        Else
                            Call CreateElements(strCustomSearch, "", rsSearch.Fields("Srch_Order_By").Value, strSql)
                        End If

                    End If

                Else
                    Call CreateElements(rsSearch.Fields("Columns1").Value, rsSearch.Fields("Headings1").Value, rsSearch.Fields("Srch_Order_By").Value, strSql)
                End If

                If Trim(m_UserName) = "NO-ID" Then
                    m_strSqlFields = RTrim(rsSearch.Fields("Srch_Order_By").Value) & " AS ID "
                Else
                    m_strSqlFields = RTrim(rsSearch.Fields("Srch_Table").Value) & ".ID AS ID "
                End If

                m_strSqlFields += ", " & m_IndexField & " "

                Call CreateSQLFieldsLine(m_strSqlFields)

            End If
            rsSearch.Close()

            '==================================================
            'Modified May 30, 2017 - T. Louzon - Add in a new AR-CUSTOMER with more fields
            'Added to 
            'cmdOpen.Visible = (_SearchElement = "AR-CUSTOMER")
            cmdOpen.Visible = (_SearchElement = "OEI-CUSTOMER")
            '==================================================

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub IndexField(ByVal pstrFields As String, ByVal pstrOrder As String)

        Dim iPos As Integer

        If Trim(pstrOrder) <> "" Then
            iPos = InStr(pstrOrder, "|")
            If iPos = 0 Then
                m_IndexField = Trim(pstrOrder)
            Else
                m_IndexField = Mid(pstrOrder, 1, iPos - 1)
            End If
        Else
            If Trim(pstrFields) <> "" Then
                iPos = InStr(pstrFields, "|")
                If iPos = 0 Then
                    m_IndexField = Trim(pstrFields)
                Else
                    m_IndexField = Mid(pstrFields, 1, iPos - 1)
                End If
            End If

        End If

    End Sub

    Private Function SetElementNameByFieldName(ByVal pstrName As Object) As String

        Dim lPos As Integer

        lPos = InStr(1, pstrName, ".")

        If lPos > 0 Then
            pstrName = RTrim(Mid(pstrName, lPos + 1)) & " "
        End If

        pstrName = Replace(pstrName, "_", " ")
        pstrName = StrConv(pstrName, VbStrConv.ProperCase)

        pstrName = Replace(pstrName, "Cus Alt Adr Cd ", "Alt Address No ")
        pstrName = Replace(pstrName, "Cus ", "Customer ")
        pstrName = Replace(pstrName, "Ord ", "Order ")
        pstrName = Replace(pstrName, "Job ", "Project ")
        pstrName = Replace(pstrName, "Desc ", "Description ")
        pstrName = Replace(pstrName, "Loc ", "Location ")
        pstrName = Replace(pstrName, "Kstdrcode ", "Cost Unit ")
        pstrName = Replace(pstrName, "Kstplcode ", "Cost Center ")
        pstrName = Replace(pstrName, "Oms25 ", "Description ")
        pstrName = Replace(pstrName, "Cmt ", "Comment ")
        pstrName = Replace(pstrName, "Id ", "ID ")

        SetElementNameByFieldName = pstrName

    End Function

    Private Sub GetSearch(Optional ByVal pstrButton As String = "", Optional ByVal pstrValue As String = "")

        Try
            m_strSqlWhereElements = CreateWhereElements(pstrButton, pstrValue)

            If m_blnAll Then
                m_strSqlSelect = "SELECT DISTINCT "
            Else
                If IsNumeric(txtTopRows.Text) Then
                    m_strSqlSelect = "SELECT TOP " & txtTopRows.Text & " WITH TIES "
                Else
                    m_strSqlSelect = "SELECT TOP 50 WITH TIES "
                End If
            End If

            m_strSqlSearchQuery = m_strSqlSelect
            m_strSqlSearchQuery = m_strSqlSearchQuery & m_strSqlFields
            m_strSqlSearchQuery = m_strSqlSearchQuery & m_strSqlFrom
            m_strSqlSearchQuery = m_strSqlSearchQuery & m_strSqlInnerJoin

            If m_strSqlSearchWhere <> "" Then
                m_strSqlSearchQuery = m_strSqlSearchQuery & " WHERE " & m_strSqlSearchWhere
            End If

            If m_strSqlWhereElements <> "" Then
                m_strSqlSearchQuery = m_strSqlSearchQuery & IIf(m_strSqlSearchWhere <> "", " AND ", " WHERE ")
                m_strSqlSearchQuery = m_strSqlSearchQuery & m_strSqlWhereElements
            End If

            m_strSqlSearchQuery = m_strSqlSearchQuery & m_strSqlOrderBy

            Dim dt As DataTable
            Dim db As New cDBA
            dt = db.DataTable(m_strSqlSearchQuery)

            dgvSearch.DataSource = dt ' rsSearch
            dgvSearch.AllowUserToAddRows = False
            dgvSearch.AllowUserToDeleteRows = False
            dgvSearch.AllowUserToOrderColumns = False

            dgvSearch.Refresh() ' .CtlRefresh()
            Call SetSearchGridHeaders()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub SetDefaultSearchElement(ByVal pstrField As String, ByVal pstrValue As String, ByVal pstrType As String)

        txtElement(0).Tag = pstrType
        txtElement(0).Text = pstrValue
        lblElement(0).Tag = pstrField
        lblElement(0).Text = pstrField

    End Sub

    Private Function CreateWhereElements(Optional ByVal pstrSearch As String = "Next", Optional ByVal pstrValue As String = "") As String

        Dim Add_SearchElements As Object
        Dim lPos As Integer
        Dim strSearch As String

        Add_SearchElements = ""

        For lPos = 0 To m_lElementCount

            If Len(RTrim(txtElement(lPos).Text)) > 0 Then
                strSearch = ""
                Select Case txtElement(lPos).Tag
                    Case CStr(2), CStr(3), CStr(4), CStr(5), CStr(7), CStr(17), CStr(22), CStr(37), CStr(131) ' numeric
                        If IsNumeric(txtElement(lPos).Text) Then
                            '                        strSearch = "(" & lblElement(lPos).Tag & " = " & _
                            ''                        txtElement(lPos).Text & ") "
                            If optStartsWith.Checked = True Then
                                strSearch = "(CAST(" & FieldName(lblElement(lPos).Tag) & " AS VARCHAR) LIKE '" & txtElement(lPos).Text & "%' ) "
                            Else
                                strSearch = "(CAST(" & FieldName(lblElement(lPos).Tag) & " AS VARCHAR) LIKE '%" & txtElement(lPos).Text & "%' ) "
                            End If
                        End If
                    Case CStr(1), CStr(129), CStr(200) ' char, varchar
                        strSearch = lblElement(lPos).Tag & " LIKE "

                        If optStartsWith.Checked = True Then
                            strSearch = "(" & FieldName(lblElement(lPos).Tag) & " LIKE '" & SqlCompliantString(Replace(txtElement(lPos).Text, " ", "% ")) & "%' OR " & FieldName(lblElement(lPos).Tag) & " LIKE '% " & SqlCompliantString(Replace(txtElement(lPos).Text, " ", "% ")) & "%'" & ") "
                        Else
                            strSearch = "(" & FieldName(lblElement(lPos).Tag) & " LIKE '%" & SqlCompliantString(Replace(txtElement(lPos).Text, " ", "% ")) & "%' " & ") "
                        End If

                    Case CStr(21), CStr(135) ' Date
                        If IsDate(txtElement(lPos).Text) Then
                            strSearch = "(" & FieldName(lblElement(lPos).Tag) & " >={d '" & Mid(txtElement(lPos).Text, 7, 4) & "-" & Mid(txtElement(lPos).Text, 4, 2) & "-" & Mid(txtElement(lPos).Text, 4, 2) & "'} ) "
                        End If
                End Select

                If Add_SearchElements = "" Then
                    Add_SearchElements = strSearch
                Else
                    Add_SearchElements = Add_SearchElements & IIf(optAnd.Checked = True, " AND ", " OR ") & strSearch
                End If

            End If

        Next lPos

        If RTrim(Add_SearchElements) <> "" Then
            Add_SearchElements = " ( " & Add_SearchElements & " ) "
        End If

        strSearch = ""
        If pstrValue <> "" Then

            'For lPos = 1 To m_lElementCount

            Select Case txtElement(1).Tag

                Case CStr(2), CStr(3), CStr(4), CStr(5), CStr(7), CStr(17), CStr(22), CStr(37), CStr(131) ' numeric
                    If IsNumeric(pstrValue) Then
                        strSearch = "(" & lblElement(1).Tag & IIf(UCase(pstrSearch) = "NEXT", ">", ">=") & pstrValue & ") "
                    End If

                Case CStr(1), CStr(129), CStr(200), CStr(4194305) ' char, varchar
                    strSearch = "(" & lblElement(1).Tag & IIf(UCase(pstrSearch) = "NEXT", ">", ">=") & "'" & SqlCompliantString(pstrValue) & "') "

                    '                Case 135 ' Date ' on ne peut pas vraiment faire de > juste sur le champ de la date...
                    '                    If IsDate(txtElement(lPos).Text) Then
                    '                        strSearch = "(" & lblElement(lPos).Tag & " >={d " & _
                    ''                        Mid(pstrValue, 7, 4) & "-" & Mid(pstrValue, 4, 2) & "-" & Mid(pstrValue, 4, 2) & "'} ) "
                    '                    End If

            End Select

        End If

        If strSearch <> "" Then
            If Add_SearchElements <> "" Then
                Add_SearchElements = Add_SearchElements & " AND " & strSearch
            Else
                Add_SearchElements = strSearch
            End If
        End If

        CreateWhereElements = Add_SearchElements

    End Function

    Private Sub SetSearchGridHeaders()

        Dim lPos As Integer

        dgvSearch.Columns(0).Visible = False
        dgvSearch.Columns(1).Visible = False

        For lPos = 1 To m_lElementCount
            dgvSearch.Columns(lPos + 1).HeaderText = lblElement(lPos).Text.ToString
            dgvSearch.Columns(lPos + 1).Width = IIf(TextWidth(lblElement(lPos).Text) > txtElement(lPos).Width, TextWidth(lblElement(lPos).Text) + 25, txtElement(lPos).Width)
        Next lPos

    End Sub

    Private Function FieldName(ByRef pstrFieldName As String) As String

        If InStr(1, pstrFieldName, " ") Then
            FieldName = Mid(pstrFieldName, 1, InStr(1, pstrFieldName, " ") - 1)
        Else
            FieldName = pstrFieldName
        End If

    End Function

    Private Function TextWidth(ByVal pstrChaine As String) As Long

        Dim g As Graphics = Me.CreateGraphics()
        Dim stringFont As New Font("Arial", 10)

        TextWidth = g.MeasureString(pstrChaine, stringFont).Width

        stringFont = Nothing
        g = Nothing

    End Function

    Public Property FoundElementIndex() As Object
        Get
            FoundElementIndex = _FoundElementIndex
        End Get
        Set(ByVal poElement As Object)
            _FoundElementIndex = poElement
        End Set
    End Property

    Public Property FoundElementValue() As Object
        Get
            FoundElementValue = _FoundElementValue
        End Get
        Set(ByVal poElement As Object)
            _FoundElementValue = poElement
        End Set
    End Property

    Public Property FoundRow() As Integer
        Get
            FoundRow = _FoundRow
        End Get
        Set(ByVal piRow As Integer)
            _FoundRow = piRow
        End Set
    End Property

    Public Property SearchElement() As String
        Get
            SearchElement = _SearchElement
        End Get
        Set(ByVal value As String)
            _SearchElement = value
        End Set
    End Property

    Public Property SearchFromWindow() As String
        Get
            SearchFromWindow = _SearchFromWindow
        End Get
        Set(ByVal value As String)
            _SearchFromWindow = value
        End Set
    End Property

    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click

        Dim lPos As Long
        For lPos = 1 To m_lElementCount
            txtElement(lPos).Text = ""
        Next lPos

    End Sub

    Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click

        Try
            If dgvSearch.Rows.Count >= 1 Then

                If dgvSearch.CurrentRow Is Nothing Then

                    _FoundElementIndex = dgvSearch.Rows(0).Cells(0).Value.ToString 'fgSearch.get_TextMatrix(fgSearch.Row, 1)
                    _FoundElementValue = dgvSearch.Rows(0).Cells(1).Value.ToString 'fgSearch.get_TextMatrix(fgSearch.Row, 2)
                    _FoundRow = 0 ' fgSearch.Row

                Else

                    _FoundElementIndex = dgvSearch.CurrentRow.Cells(0).Value.ToString 'fgSearch.get_TextMatrix(fgSearch.Row, 1)
                    _FoundElementValue = dgvSearch.CurrentRow.Cells(1).Value.ToString 'fgSearch.get_TextMatrix(fgSearch.Row, 2)
                    _FoundRow = dgvSearch.CurrentRow.Index ' fgSearch.Row

                End If

            Else

                _FoundElementIndex = ""
                _FoundElementValue = ""
                _FoundRow = -1

            End If

            Me.Close()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    Public ReadOnly Property DataGrid() As DataGridView
        Get
            DataGrid = dgvSearch
        End Get
    End Property


    Public Function Connection() As ADODB.Connection

        Dim cnConnection As String = g_User.Conn_String

        Dim Conn As New ADODB.Connection
        Conn.ConnectionString = cnConnection

        Conn.Open()
        Conn.Close()

        Connection = Conn

    End Function

    ''Private Sub fgSearch_KeyDownEvent(ByVal sender As Object, ByVal e As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles fgSearch.KeyDownEvent
    'Private Sub fgSearch_KeyDownEvent(ByVal sender As Object 'ByVal e As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent)

    '    Try
    '        If e.keyCode = Keys.Return Then ' e.Control And e.KeyValue = Keys.F Then
    '            cmdSearch.PerformClick()
    '        End If

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Private Sub cmdOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpen.Click

        Try
            If dgvSearch.Rows.Count = 0 Then Exit Sub

            Dim oCustomer As cCustomerInfo

            If dgvSearch.CurrentRow Is Nothing Then
                oCustomer = New cCustomerInfo(dgvSearch.Rows(0).Cells(1).Value.ToString)
            Else
                oCustomer = New cCustomerInfo(dgvSearch.CurrentRow.Cells(1).Value.ToString)
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub frmSearch_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        optContains.Checked = True

    End Sub

    Private Sub dgvSearch_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSearch.DoubleClick

        cmdSelect.PerformClick()

    End Sub


    Private Sub dgvSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvSearch.KeyDown

        Try
            If e.KeyCode = Keys.Return Then ' e.Control And e.KeyValue = Keys.F Then
                cmdSelect.PerformClick()
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function CustomOrder(ByVal pstrSubKey As String) As String

        CustomOrder = ""

    End Function

    Private Sub txtTopRows_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTopRows.KeyDown

        Select Case e.KeyCode
            Case Keys.Return
                cmdSearch.PerformClick()
        End Select

    End Sub

    Private Sub dgvSearch_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSearch.CellContentClick

    End Sub
End Class

