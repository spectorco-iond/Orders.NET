Option Strict Off
Option Explicit On
Friend Class ucTaxes
	Inherits System.Windows.Forms.UserControl

    'Private m_oOrder As New cOrder
    Private blnLoading As Boolean = False

    Private blnSetTags As Boolean = False
    Private blnForceFocus As Boolean = False


#Region "Private elements control functions ###############################"

    ' Returns the label to show complimentary data on.
    Private Function GetDescriptionLabelByElement(ByVal txtElementName As String) As Label

        GetDescriptionLabelByElement = Nothing

        Try
            Select Case txtElementName

                Case "txtTax_Cd"
                    GetDescriptionLabelByElement = lblTax_Cd_Desc

                Case "txtTax_Cd_2"
                    GetDescriptionLabelByElement = lblTax_Cd_2_Desc

                Case "txtTax_Cd_3"
                    GetDescriptionLabelByElement = lblTax_Cd_3_Desc

                Case "txtTax_Sched"
                    GetDescriptionLabelByElement = lblTax_Sched_Desc

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Returns the textbox or the combobox to apply the search from the search button
    Private Function GetElementByControlName(ByVal txtElement As String) As Object

        GetElementByControlName = ""

        Try
            Select Case txtElement

                Case "txtTax_Cd", "cmdTax_Cd_Search", "cmdTax_Cd_Notes"
                    GetElementByControlName = txtTax_Cd

                Case "txtTax_Cd_2", "cmdTax_Cd_2_Search", "cmdTax_Cd_2_Notes"
                    GetElementByControlName = txtTax_Cd_2

                Case "txtTax_Cd_3", "cmdTax_Cd_3_Search", "cmdTax_Cd_3_Notes"
                    GetElementByControlName = txtTax_Cd_3

                Case "txtTax_Sched", "cmdTax_Sched_Search", "cmdTax_Sched_Notes"
                    GetElementByControlName = txtTax_Sched

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

#End Region


#Region "Private events - GotFocus ########################################"

    ' Logical sequence get focus to force focus on needed fields
    Private Sub Element_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        chkTax_Fg.GotFocus, txtTax_Sched.GotFocus, cmdTax_Sched_Search.GotFocus, cmdTax_Sched_Notes.GotFocus, _
        txtTax_Cd.GotFocus, cmdTax_Cd_Search.GotFocus, cmdTax_Cd_Notes.GotFocus, txtTax_Pct.GotFocus, _
        txtTax_Cd_2.GotFocus, cmdTax_Cd_2_Search.GotFocus, cmdTax_Cd_2_Notes.GotFocus, txtTax_Pct_2.GotFocus, _
        txtTax_Cd_3.GotFocus, cmdTax_Cd_3_Search.GotFocus, cmdTax_Cd_3_Notes.GotFocus, txtTax_Pct_3.GotFocus

        Dim lTabIndex As Integer

        Try
            If TypeOf sender Is xTextBox Then

                Dim txtElement As New xTextBox
                txtElement = DirectCast(sender, xTextBox)
                lTabIndex = txtElement.TabIndex

            ElseIf TypeOf sender Is Button Then

                Dim txtElement As New Button
                txtElement = DirectCast(sender, Button)
                lTabIndex = txtElement.TabIndex

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        blnForceFocus = ForceFocus(lTabIndex)

    End Sub

#End Region


#Region "Private events - LostFocus #######################################"

    ' Sets field color to white or red when data not found in DB
    Private Sub Element_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
                    txtTax_Cd.LostFocus, txtTax_Cd_2.LostFocus, txtTax_Cd_3.LostFocus

        ' Les instructions et les commentaires ne generent pas d'erreurs

        Dim txtElement As New Object

        Try
            If TypeOf sender Is xTextBox Then

                txtElement = New xTextBox
                txtElement = DirectCast(sender, xTextBox)

            ElseIf TypeOf sender Is ComboBox Then

                txtElement = New ComboBox
                txtElement = DirectCast(sender, ComboBox)

            End If

            If ttTooltip.GetToolTip(txtElement) = "Record not on file." Then
                txtElement.BackColor = Color.FromArgb(255, 192, 192)
                'ForceFocus(txtElement)
            Else
                txtElement.BackColor = Color.White
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Private events - KeyDown #########################################"

    Private Sub Pressed_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
        txtTax_Cd.KeyDown, txtTax_Cd_2.KeyDown, txtTax_Cd_3.KeyDown

        Dim txtElement As New Object

        Try
            If TypeOf sender Is xTextBox Then

                txtElement = New xTextBox
                txtElement = DirectCast(sender, xTextBox)

            ElseIf TypeOf sender Is ComboBox Then

                txtElement = New ComboBox
                txtElement = DirectCast(sender, ComboBox)

            End If

            Select Case e.KeyValue

                ' Save changes ?
                Case Keys.Escape
                    'Dim oReponse As New MsgBoxResult
                    'oReponse = MsgBox("Save Changes ?", MsgBoxStyle.YesNoCancel, "OE Interface")
                    'If oReponse = MsgBoxResult.Yes Then
                    '    MsgBox("On save")
                    'ElseIf oReponse = MsgBoxResult.No Then
                    '    MsgBox("on annule la commande")
                    'Else
                    '    MsgBox("on fait rien")
                    'End If

                    ' Search
                Case Keys.F7
                    Dim oButton As Button
                    oButton = GetSearchControlByControlName(sender.name)
                    If Not (oButton.Name Is Nothing) Then oButton.PerformClick()

                    ' Notes
                Case Keys.F8
                    Dim oButton As Button
                    oButton = GetNotesControlByControlName(sender.name)
                    If Not (oButton.Name Is Nothing) Then oButton.PerformClick()

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Private events - TextChanged #####################################"

    ' Autocomplete fields when needed, change fore colors and fill complementary labels
    Private Sub Element_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
                txtTax_Cd.TextChanged, txtTax_Cd_2.TextChanged, txtTax_Cd_3.TextChanged, txtTax_Sched.TextChanged

        Try
            Dim lTabIndex As Integer

            Dim txtElement As New Object
            Dim lblDescriptionLabel As New Label

            If TypeOf sender Is xTextBox Then
                'Dim a1 As TextBox
                txtElement = New xTextBox
                txtElement = DirectCast(sender, xTextBox)
                lTabIndex = txtElement.TabIndex
                'Attempted to read or write protected memory. This is often an indication that other memory is corrupt.

                Dim txtAutoComplete As New xTextBox
                txtAutoComplete = DirectCast(sender, xTextBox)

            ElseIf TypeOf sender Is ComboBox Then

                txtElement = New ComboBox
                txtElement = DirectCast(sender, ComboBox)
                lTabIndex = txtElement.TabIndex

            End If

            lblDescriptionLabel = GetDescriptionLabelByElement(txtElement.name)

            If RTrim(Trim(txtElement.Text)) = String.Empty Then
                If Not (lblDescriptionLabel Is Nothing) Then
                    lblDescriptionLabel.Text = ""
                End If
                ttTooltip.SetToolTip(txtElement, "")
                If txtElement.backcolor = Color.FromArgb(255, 192, 192) Then
                    txtElement.backcolor = Color.White
                End If
                Exit Sub
            End If

            Dim oSearch As New cSearch()
            oSearch.SearchElement = GetSearchElementByControlName(txtElement.name)
            '' oSearch.SetDefaultSearchElement(GetSearchFieldByElement(oSearch.SearchElement), RTrim(txtElement.Text), GetSearchTypeByControlName(txtElement))
            oSearch.SetDefaultSearchElement(GetSearchFieldByElement(oSearch.SearchElement), RTrim(txtElement.Text.ToString.ToUpper), "129")
            If Trim(RTrim(txtElement.text)).Length > 2 Then
                'If Trim(RTrim(txtElement.text)).Length = 3 Then ' We load search only at 3
                oSearch.GetResults(True)
            Else
                oSearch.GetResults()
            End If

            If oSearch.Form.dgvSearch.Rows.Count > 2 And Trim(RTrim(sender.text.ToString.ToUpper)) <> Trim(RTrim(oSearch.Form.FoundElementValue)) Then
                If Trim(RTrim(txtElement.text)).Length <= 3 Then
                    If TypeOf sender Is xTextBox Then
                        'Dim acscAutoComplete As New AutoCompleteStringCollection()
                        'For lPos As Integer = 1 To oSearch.Form.fgSearch.Rows - 1
                        '    acscAutoComplete.Add(oSearch.Form.TextMatrix(lPos, 2))
                        '    'acscAutoComplete.Add("Salut" & lPos)
                        'Next lPos

                        'SyncLock txtElement.AutoCompleteCustomSource.SyncRoot
                        '    txtElement.AutoCompleteCustomSource = acscAutoComplete
                        'End SyncLock
                        'acscAutoComplete = Nothing
                    Else
                        'ttToolTip.SetToolTip(txtElement, "Record not on file.")
                    End If
                    ttTooltip.SetToolTip(txtElement, "Record not on file.")
                End If
            ElseIf oSearch.Form.FoundRow = 1 And Trim(RTrim(sender.text.ToString.ToUpper)) = Trim(RTrim(oSearch.Form.FoundElementValue)) Then ' oSearch.Form.fgSearch.Rows = 2 And oSearch.Form.FoundRow = 1 Then
                ' Be careful to always check if the value is the same, else we will enter in a deadlock
                'If oSearch.Form.FoundElementValue <> txtElement.text Then
                '    txtElement.text = oSearch.Form.FoundElementValue
                'End If
                'txtElement.AutoCompleteCustomSource = Nothing
                ttTooltip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(oSearch.Form.FoundRow).Cells(3).Value)
                If Not (lblDescriptionLabel Is Nothing) Then
                    lblDescriptionLabel.Text = oSearch.Form.DataGrid.Rows(oSearch.Form.FoundRow).Cells(3).Value
                End If
                txtElement.backcolor = Color.White
            Else
                ttTooltip.SetToolTip(txtElement, "Record not on file.")
                'txtElement.backcolor = Color.FromArgb(255, 192, 192)
                If Not (lblDescriptionLabel Is Nothing) Then
                    lblDescriptionLabel.Text = ""
                End If
                'Dim acscAutoComplete As New AutoCompleteStringCollection()
                'SyncLock txtElement.AutoCompleteCustomSource.SyncRoot
                '    txtElement.AutoCompleteCustomSource = acscAutoComplete
                'End SyncLock
                'acscAutoComplete = Nothing
            End If

            oSearch = Nothing

            'If blnSetTags Then txtElement.Tag = txtElement.Text.ToString.ToUpper
            If blnSetTags Then txtElement.OldValue = txtElement.Text.ToString.ToUpper

            'txtElement.forecolor = IIf(txtElement.tag <> txtElement.text.ToString.ToUpper And Not blnLoading, Color.Blue, Color.Black)
            txtElement.forecolor = IIf(txtElement.OldValue <> txtElement.text.ToString.ToUpper And Not blnLoading, Color.Blue, Color.Black)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Private Notes control functions ##################################"

    ' Get Field_ID for SEARCHES_SQL table for a control name
    Private Function GetNotesElementByControlName(ByVal txtElement As String) As String

        GetNotesElementByControlName = ""

        Try
            Select Case txtElement

                Case "cmdTax_Cd_Notes", "cmdTax_Cd_Notes_2", "cmdTax_Cd_Notes_3"
                    GetNotesElementByControlName = "SY-TAX_CODE"

                Case "cmdTax_Sched_Notes"
                    GetNotesElementByControlName = "SY-TAX-SCHED"

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Returns the search button control associated with a textbox or a combobox
    Private Function GetNotesControlByControlName(ByVal txtElement As String) As Button

        GetNotesControlByControlName = New Button

        Try
            Select Case txtElement

                Case "txtTax_Cd"
                    GetNotesControlByControlName = cmdTax_Cd_Notes

                Case "txtTax_Cd_2"
                    GetNotesControlByControlName = cmdTax_Cd_2_Notes

                Case "txtTax_Cd_3"
                    GetNotesControlByControlName = cmdTax_Cd_3_Notes

                Case "txtTax_Sched"
                    GetNotesControlByControlName = cmdTax_Sched_Notes

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Search buttons click - open search window for element
    Private Sub Notes_Element_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _
        cmdTax_Cd_Notes.Click, cmdTax_Cd_2_Notes.Click, cmdTax_Cd_3_Notes.Click, cmdTax_Sched_Notes.Click

        Dim btnElement As Button

        Try

            If TypeOf eventSender Is Button Then

                btnElement = New Button
                btnElement = DirectCast(eventSender, Button)

                Dim txtElement As Object
                txtElement = GetElementByControlName(btnElement.Name)

                Dim oNotes As New CNotes(GetNotesElementByControlName(btnElement.Name), txtElement.Text.ToString)
                oNotes = Nothing

                txtElement.focus()

            End If
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Private Search control functions #################################"

    ' Search buttons click - open search window for element
    Private Sub Search_Element_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _
                cmdTax_Cd_Search.Click, cmdTax_Cd_2_Search.Click, cmdTax_Cd_3_Search.Click, cmdTax_Sched_Search.Click

        Dim btnElement As Button

        Try

            If TypeOf eventSender Is Button Then

                btnElement = New Button
                btnElement = DirectCast(eventSender, Button)

                Dim oSearch As New cSearch(GetSearchElementByControlName(btnElement.Name))
                If oSearch.Form.FoundRow <> -1 Then

                    Dim txtElement As Object

                    txtElement = GetElementByControlName(btnElement.Name)
                    txtElement.Text = oSearch.Form.FoundElementValue
                    ttTooltip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(oSearch.Form.FoundRow).Cells(3).Value)

                    txtElement.focus()

                End If
                oSearch = Nothing

            End If
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Returns the search button control associated with a textbox or a combobox
    Private Function GetSearchControlByControlName(ByVal txtElement As String) As Button

        GetSearchControlByControlName = New Button

        Try
            Select Case txtElement

                Case "txtTax_Cd"
                    GetSearchControlByControlName = cmdTax_Cd_Search

                Case "txtTax_Cd_2"
                    GetSearchControlByControlName = cmdTax_Cd_2_Search

                Case "txtTax_Cd_3"
                    GetSearchControlByControlName = cmdTax_Cd_3_Search

                Case "txtTax_Sched"
                    GetSearchControlByControlName = cmdTax_Sched_Search

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Get Field_ID for SEARCHES_SQL table for a control name
    Private Function GetSearchElementByControlName(ByVal txtElement As String) As String

        GetSearchElementByControlName = ""

        Try
            Select Case txtElement

                Case "txtTax_Cd", "txtTax_Cd_2", "txtTax_Cd_3", "cmdTax_Cd_Search", "cmdTax_Cd_2_Search", "cmdTax_Cd_3_Search"
                    GetSearchElementByControlName = "SY-TAX-CODE"

                Case "txtTax_Sched", "cmdTax_Sched_Search"
                    GetSearchElementByControlName = "SY-TAX-SCHED"

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Get search field from SEARCHES_SQL for a control name (to link with FieldID)
    Private Function GetSearchFieldByElement(ByVal pstrSearchElement As String) As String

        GetSearchFieldByElement = ""

        Try
            Select Case pstrSearchElement

                Case "SY-TAX-CODE"
                    GetSearchFieldByElement = "TAXDETL_SQL.tax_cd"

                Case "SY-TAX-SCHED"
                    GetSearchFieldByElement = "TAXSCHED_SQL.tax_sched"

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

#End Region


#Region "Private Z other functions ########################################"

    Private Sub ForceFocus(ByRef oElement As Object)

        Dim txtElement As New Object

        Try
            If TypeOf oElement Is xTextBox Then

                txtElement = New xTextBox
                txtElement = DirectCast(oElement, xTextBox)

            ElseIf TypeOf oElement Is ComboBox Then

                txtElement = New ComboBox
                txtElement = DirectCast(oElement, ComboBox)

            End If

            txtElement.focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' This proc checks the tabindex of the caller and compare it to needed fields.
    ' Also, completes some fields depending on position from tabindex order.
    Private Function ForceFocus(ByVal piTabIndex As Integer) As Boolean

        ForceFocus = False

        Try
            If piTabIndex >= txtTax_Cd.TabIndex And Trim(txtTax_Sched.Text) <> "" Then
                txtTax_Sched.Focus()
                ForceFocus = True
                Exit Function
            End If

            If piTabIndex = txtTax_Sched.TabIndex And Trim(txtTax_Cd.Text) <> "" Then
                txtTax_Cd.Focus()
                ForceFocus = True
                Exit Function
            End If

            If piTabIndex > txtTax_Cd.TabIndex And Trim(txtTax_Cd.Text) = "" Then
                chkTax_Fg.Focus()
                ForceFocus = True
                Exit Function
            End If

            If piTabIndex >= txtTax_Cd_2.TabIndex And Trim(txtTax_Cd.Text) = "" Then
                txtTax_Cd.Focus()
                ForceFocus = True
                Exit Function
            End If

            If piTabIndex >= txtTax_Pct_2.TabIndex And Trim(txtTax_Cd_2.Text) = "" Then
                chkTax_Fg.Focus()
                ForceFocus = True
                Exit Function
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

#End Region


#Region "Public order management procedures ###############################"

    Public Sub Fill() ' ByRef m_oOrder As cOrder)

        blnLoading = True

        'm_oOrder = m_oOrder

        If Trim(m_oOrder.Ordhead.Ord_No).Length <> 0 Then
            chkTax_Fg.Checked = (m_oOrder.Ordhead.Tax_Fg = "Y")
            txtTax_Sched.Text = m_oOrder.Ordhead.Tax_Sched
            lblTax_Cd_Desc.Text = m_oOrder.Descriptions.TaxSchedule

            txtTax_Cd.Text = m_oOrder.Ordhead.Tax_Cd
            txtTax_Pct.Text = (m_oOrder.Ordhead.Tax_Pct) ' String.Format("{0:n4}", m_oOrder.Ordhead.Tax_Pct)
            lblTax_Cd_Desc.Text = m_oOrder.Descriptions.TaxCode1

            txtTax_Cd_2.Text = m_oOrder.Ordhead.Tax_Cd_2
            txtTax_Pct_2.Text = (m_oOrder.Ordhead.Tax_Pct_2) ' String.Format("{0:n4}", m_oOrder.Ordhead.Tax_Pct_2)
            lblTax_Cd_2_Desc.Text = m_oOrder.Descriptions.TaxCode2

            txtTax_Cd_3.Text = m_oOrder.Ordhead.Tax_Cd_3
            txtTax_Pct_3.Text = (m_oOrder.Ordhead.Tax_Pct_3) ' String.Format("{0:n4}", m_oOrder.Ordhead.Tax_Pct_3)
            lblTax_Cd_3_Desc.Text = m_oOrder.Descriptions.TaxCode3
        Else
            Reset() ' m_oOrder)
        End If

        blnLoading = False

    End Sub

    Public Sub Reset() ' ByRef pOrder As cOrder)

        chkTax_Fg.Text = (m_oOrder.Ordhead.Tax_Fg = "Y")
        txtTax_Sched.Text = m_oOrder.Ordhead.Tax_Sched
        txtTax_Cd.Text = m_oOrder.Ordhead.Tax_Cd
        txtTax_Pct.Text = m_oOrder.Ordhead.Tax_Pct ' String.Format("{0:n4}", "0")
        lblTax_Cd_Desc.Text = m_oOrder.Descriptions.TaxCode1
        txtTax_Cd_2.Text = m_oOrder.Ordhead.Tax_Cd_2
        txtTax_Pct_2.Text = m_oOrder.Ordhead.Tax_Pct_2 ' String.Format("{0:n4}", "0")
        lblTax_Cd_2_Desc.Text = m_oOrder.Descriptions.TaxCode2
        txtTax_Cd_3.Text = m_oOrder.Ordhead.Tax_Cd_3
        txtTax_Pct_3.Text = m_oOrder.Ordhead.Tax_Pct_3 ' String.Format("{0:n4}", "0")
        lblTax_Cd_3_Desc.Text = m_oOrder.Descriptions.TaxCode3

    End Sub

    Public Sub Save()

        Exit Sub

        ' Do nothing here for now
        m_oOrder.Ordhead.Tax_Fg = chkTax_Fg.Text
        m_oOrder.Ordhead.Tax_Sched = txtTax_Sched.Text
        m_oOrder.Ordhead.Tax_Cd = txtTax_Cd.Text
        m_oOrder.Ordhead.Tax_Pct = txtTax_Pct.Text
        m_oOrder.Ordhead.Tax_Cd_2 = txtTax_Cd_2.Text
        m_oOrder.Ordhead.Tax_Pct_2 = txtTax_Pct_2.Text
        m_oOrder.Ordhead.Tax_Cd_3 = txtTax_Cd_3.Text
        m_oOrder.Ordhead.Tax_Pct_3 = txtTax_Pct_3.Text

    End Sub

#End Region


#Region "Public properties ################################################"

    Public ReadOnly Property Order() As cOrder
        Get
            Order = m_oOrder
        End Get
    End Property

#End Region


    'Private Sub ucTaxes_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

    '    ' If Ctl-Number, sets the new tab.
    '    If e.Control Then
    '        Select Case e.KeyCode
    '            Case Keys.D1
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Header)
    '            Case Keys.D2
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CustomerContact)
    '            Case Keys.D3
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Lines)
    '            Case Keys.D4
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Documents)
    '            Case Keys.D5
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.HeaderAll)
    '            Case Keys.D6
    '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Taxes)
    '            Case Keys.D7
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Salesperson)
    '            Case Keys.D8
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
    '            Case Keys.D9
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
    '        End Select

    '    End If

    'End Sub

End Class