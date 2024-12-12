Option Strict Off
Option Explicit On

Friend Class ucSalesperson
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

                'Case "txtOrder"
                '    GetSearchElementByControlName = "OE-ORDER-NO"

                Case "txtSlspsn_No"
                    GetDescriptionLabelByElement = lblSlspsn_No_Desc

                Case "txtSlspsn_No_2"
                    GetDescriptionLabelByElement = lblSlspsn_No_2_Desc

                Case "txtSlspsn_No_3"
                    GetDescriptionLabelByElement = lblSlspsn_No_3_Desc

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

                Case "cmdSlspsn_No", "cmdSlspsn_No_Notes"
                    GetElementByControlName = txtSlspsn_No

                Case "cmdSlspsn_No_2", "cmdSlspsn_No_Notes_2"
                    GetElementByControlName = txtSlspsn_No_2

                Case "cmdSlspsn_No_3", "cmdSlspsn_No_Notes_3"
                    GetElementByControlName = txtSlspsn_No_3

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

#End Region


#Region "Private events - LostFocus #######################################"

    ' Sets field color to white or red when data not found in DB
    Private Sub Element_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
                    txtSlspsn_No.LostFocus, txtSlspsn_No_2.LostFocus, txtSlspsn_No_3.LostFocus

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

            If ttToolTip.GetToolTip(txtElement) = "Record not on file." Then
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
        txtSlspsn_No.KeyDown, txtSlspsn_No_2.KeyDown, txtSlspsn_No_3.KeyDown

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
                txtSlspsn_No.TextChanged, txtSlspsn_No_2.TextChanged, txtSlspsn_No_3.TextChanged

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
                ttToolTip.SetToolTip(txtElement, "")
                If txtElement.backcolor = Color.FromArgb(255, 192, 192) Then
                    txtElement.backcolor = Color.White
                End If
                Exit Sub
            End If

            If Not (txtElement.name = "txtCus_Alt_Adr_Cd" And txtElement.text = "SAME") Then
                Dim oSearch As New cSearch()
                oSearch.SearchElement = GetSearchElementByControlName(txtElement.name)
                '' oSearch.SetDefaultSearchElement(GetSearchFieldByElement(oSearch.SearchElement), RTrim(txtElement.Text), GetSearchTypeByControlName(txtElement))
                oSearch.SetDefaultSearchElement(GetSearchFieldByElement(oSearch.SearchElement), Trim(txtElement.Text.ToString.ToUpper), "129")
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
                        ttToolTip.SetToolTip(txtElement, "Record not on file.")
                    End If
                ElseIf oSearch.Form.FoundRow = 1 And Trim(RTrim(sender.text.ToString.ToUpper)) = Trim(RTrim(oSearch.Form.FoundElementValue)) Then ' oSearch.Form.fgSearch.Rows = 2 And oSearch.Form.FoundRow = 1 Then
                    ' Be careful to always check if the value is the same, else we will enter in a deadlock
                    'If oSearch.Form.FoundElementValue <> txtElement.text Then
                    '    txtElement.text = oSearch.Form.FoundElementValue
                    'End If
                    'txtElement.AutoCompleteCustomSource = Nothing
                    ttToolTip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(oSearch.Form.FoundRow).Cells(3).Value)
                    If Not (lblDescriptionLabel Is Nothing) Then
                        lblDescriptionLabel.Text = oSearch.Form.DataGrid.Rows(oSearch.Form.FoundRow).Cells(3).Value
                    End If
                    txtElement.backcolor = Color.White
                Else
                    ttToolTip.SetToolTip(txtElement, "Record not on file.")
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

            End If

            'If blnSetTags Then txtElement.Tag = txtElement.Text.ToString.ToUpper
            If blnSetTags Then txtElement.OldValue = txtElement.Text.ToString.ToUpper

            'txtElement.forecolor = IIf(txtElement.tag <> txtElement.text.ToString.ToUpper And Not blnLoading, Color.Blue, Color.Black)
            txtElement.forecolor = IIf(txtElement.OldValue <> txtElement.text.ToString.ToUpper And Not blnLoading, Color.Blue, Color.Black)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtSlspsn_No_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        m_oOrder.Descriptions.Salesperson1 = m_oOrder.Ordhead.Get_HumRes_ID_Description(m_oOrder.Ordhead.Slspsn_No)
        lblSlspsn_No_Desc.Text = m_oOrder.Descriptions.Salesperson1
    End Sub

#End Region


#Region "Private Notes control functions ##################################"

    ' Get Field_ID for SEARCHES_SQL table for a control name
    Private Function GetNotesElementByControlName(ByVal txtElement As String) As String

        GetNotesElementByControlName = ""

        Try
            Select Case txtElement

                Case "cmdSlspsn_No_Notes", "cmdSlspsn_No_Notes_2", "cmdSlspsn_No_Notes_3"
                    GetNotesElementByControlName = "AR-SALESPERSON"

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

                Case "txtSlspsn_No"
                    GetNotesControlByControlName = cmdSlspsn_No_Notes

                Case "txtSlspsn_No_2"
                    GetNotesControlByControlName = cmdSlspsn_No_Notes_2

                Case "txtSlspsn_No_3"
                    GetNotesControlByControlName = cmdSlspsn_No_Notes_3

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Search buttons click - open search window for element
    Private Sub Notes_Element_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _
        cmdSlspsn_No_Notes.Click, cmdSlspsn_No_Notes_2.Click, cmdSlspsn_No_Notes_3.Click

        Dim btnElement As Button

        Try

            If TypeOf eventSender Is Button Then

                btnElement = New Button
                btnElement = DirectCast(eventSender, Button)

                Dim txtElement As Object
                txtElement = GetElementByControlName(btnElement.Name)

                Dim oNotes As New CNotes(GetNotesElementByControlName(btnElement.Name), txtElement.Text.ToString)
                oNotes = Nothing

                txtElement.Focus()

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Private Search control functions #################################"

    ' Search buttons click - open search window for element
    Private Sub Search_Element_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _
                cmdSlspsn_No.Click, cmdSlspsn_No_2.Click, cmdSlspsn_No_3.Click

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
                    ttToolTip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(oSearch.Form.FoundRow).Cells(3).Value)

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

                Case "txtSlspsn_No"
                    GetSearchControlByControlName = cmdSlspsn_No

                Case "txtSlspsn_No_2"
                    GetSearchControlByControlName = cmdSlspsn_No_2

                Case "txtSlspsn_No_3"
                    GetSearchControlByControlName = cmdSlspsn_No_3

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

                Case "txtSlspsn_No", "txtSlspsn_No_2", "txtSlspsn_No_3", "cmdSlspsn_No", "cmdSlspsn_No_2", "cmdSlspsn_No_3"
                    GetSearchElementByControlName = "AR-SALESPERSON"

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

                Case "AR-SALESPERSON"
                    GetSearchFieldByElement = "arslmfil_sql.humres_id"

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

#End Region


#Region "Public order management procedures ###############################"

    Public Sub Fill() ' ByRef m_oOrder As cOrder)

        blnLoading = True

        ' m_oOrder = m_oOrder

        If Trim(m_oOrder.Ordhead.Ord_No).Length <> 0 Then
            txtSlspsn_No.Text = m_oOrder.Ordhead.Slspsn_No
            lblSlspsn_No_Desc.Text = m_oOrder.Descriptions.Salesperson1
            txtSlspsn_No_2.Text = m_oOrder.Ordhead.Slspsn_No_2
            lblSlspsn_No_2_Desc.Text = m_oOrder.Descriptions.Salesperson2
            txtSlspsn_No_3.Text = m_oOrder.Ordhead.Slspsn_No_3
            lblSlspsn_No_3_Desc.Text = m_oOrder.Descriptions.Salesperson3

            txtSlspsn_Pct_Comm.Text = String.Format("{0:n3}", m_oOrder.Ordhead.Slspsn_Pct_Comm)
            txtSlspsn_Pct_Comm_2.Text = String.Format("{0:n3}", m_oOrder.Ordhead.Slspsn_Pct_Comm_2)
            txtSlspsn_Pct_Comm_3.Text = String.Format("{0:n3}", m_oOrder.Ordhead.Slspsn_Pct_Comm_3)
        Else
            Reset() ' m_oOrder)
        End If

        blnLoading = False

    End Sub

    Public Sub Reset() ' ByRef pOrder As cOrder)

        txtSlspsn_No.Text = ""
        txtSlspsn_No_2.Text = ""
        txtSlspsn_No_3.Text = ""

        lblSlspsn_No_Desc.Text = ""
        lblSlspsn_No_2_Desc.Text = ""
        lblSlspsn_No_3_Desc.Text = ""

        txtSlspsn_Pct_Comm.Text = ""
        txtSlspsn_Pct_Comm_2.Text = ""
        txtSlspsn_Pct_Comm_3.Text = ""

    End Sub

    Public Sub Save()

        Exit Sub

        ' Do nothing here for now
        m_oOrder.Ordhead.Slspsn_No = txtSlspsn_No.Text
        m_oOrder.Ordhead.Slspsn_No_2 = txtSlspsn_No_2.Text
        m_oOrder.Ordhead.Slspsn_No_3 = txtSlspsn_No_3.Text

        m_oOrder.Ordhead.Slspsn_Pct_Comm = txtSlspsn_Pct_Comm.Text
        m_oOrder.Ordhead.Slspsn_Pct_Comm_2 = txtSlspsn_Pct_Comm_2.Text
        m_oOrder.Ordhead.Slspsn_Pct_Comm_3 = txtSlspsn_Pct_Comm_3.Text

    End Sub

#End Region


#Region "Public properties ################################################"

    Public ReadOnly Property Order() As cOrder
        Get
            Order = m_oOrder
        End Get
    End Property

#End Region


    'Private Sub ucSalesperson_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

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
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Taxes)
    '            Case Keys.D7
    '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Salesperson)
    '            Case Keys.D8
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
    '            Case Keys.D9
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
    '        End Select

    '    End If

    'End Sub

End Class