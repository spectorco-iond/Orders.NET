Option Strict Off
Option Explicit On

Imports Microsoft.VisualBasic.PowerPacks
Imports System.IO
Imports System.Data.SqlClient

Friend Class ucOrder

    Inherits System.Windows.Forms.UserControl

    Private _Item_No As String
    Private blnInvalidate As Boolean = False
    Private iRowPos As Integer = 0
    Private iColPos As Integer = 0
    Private blnLoading As Boolean = False

    Private blnForceFocus As Boolean = False
    Private m_DateButton As Button

    Public dsDataSet As New DataSet()

    Private blnSetTags As Boolean = False

    ' Order line column positions
    Public Enum _Columns
        _Item = 0
        _Loc
        _QtyOrdered
        _QtyShipped
        _UnitPrice
        _DiscPct
        _ExtPrice
        _DateRequested
        _DatePromised
        _ReqShip
        _Description
        _Description2
        _UOM
        _TaxSchd
        _ECMRevision
        _CreateModifyKit
        _NoteText
        _BackorderItem
        _Selected
        _Recalc
        _LineComments
        _Route
    End Enum

    ' Search buttons click - open search window for element
    Private Sub Element_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _
                cmdBillToCountrySearch.Click, cmdApplyToSearch.Click, _
                cmdCommentsSearch.Click, cmdCostCenterSearch.Click, cmdCustomerSearch.Click, _
                cmdInstructionsSearch.Click, cmdLocationSearch.Click, cmdOrderSearch.Click, _
                cmdProfitCenterSearch.Click, cmdProjectSearch.Click, cmdSalespersonSearch.Click, _
                cmdShipToCountrySearch.Click, cmdShipViaSearch.Click, cmdTermsSearch.Click

        If blnForceFocus Then Exit Sub

        Dim btnElement As Button

        If TypeOf eventSender Is Button Then

            btnElement = New Button
            btnElement = DirectCast(eventSender, Button)

            Dim oSearch As New cSearch(GetSearchElementByControlName(btnElement.Name))
            If oSearch.Form.FoundRow <> -1 Then
                Dim txtElement As Object
                txtElement = GetElementByControlName(btnElement.Name)
                txtElement.Text = oSearch.Form.FoundElementValue
                ttToolTip.SetToolTip(txtElement, oSearch.Form.TextMatrix(oSearch.Form.FoundRow, 3))
            End If
            oSearch = Nothing

            If btnElement.Name = "cmdCustomerSearch" Then txtCustomer_Validated(txtCustomer, New System.EventArgs)

        End If

    End Sub

    ' Date buttons click - open DateControl for element
    Private Sub Element_DateButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDateTP.Click, cmdShipDateTP.Click

        If blnForceFocus Then Exit Sub

        m_DateButton = New Button
        m_DateButton = DirectCast(sender, Button)

        mcCalendar.Top = m_DateButton.Top
        mcCalendar.Left = m_DateButton.Left
        mcCalendar.Visible = True
        mcCalendar.SetDate(Date.Now)

        Select Case m_DateButton.Name
            Case "cmdDateTP"
                If IsDate(txtOrderDate.Text) Then mcCalendar.SetDate(txtOrderDate.Text)

            Case "cmdShipDateTP"
                If IsDate(txtShipDate.Text) Then mcCalendar.SetDate(txtShipDate.Text)

        End Select

    End Sub

    ' Logical sequence get focus to force focus on needed fields
    Private Sub Element_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
                    cmdOrderNotes.GotFocus, txtOrderDate.GotFocus, cmdDateTP.GotFocus, txtApplyTo.GotFocus, _
                    cmdApplyToNotes.GotFocus, txtShipDate.GotFocus, cmdShipDateTP.GotFocus, txtShipTo.GotFocus, _
                    txtCustomer.GotFocus, cmdCustomerNotes.GotFocus, cboShipTo.GotFocus, cmdShipToNotes.GotFocus, _
                    txtPONumber.GotFocus, txtSalesPerson.GotFocus, cmdSalespersonNotes.GotFocus, txtProject.GotFocus, _
                    cmdProjectNotes.GotFocus, txtInstructions.GotFocus, txtLocation.GotFocus, cmdLocationNotes.GotFocus, _
                    cboShipVia.GotFocus, cmdShipViaNotes.GotFocus, txtTerms.GotFocus, cmdTermsNotes.GotFocus, _
                    txtDiscPct.GotFocus, txtProfitCenter.GotFocus, cmdProfitCenterNotes.GotFocus, _
                    txtCostCenter.GotFocus, cmdCostCenterNotes.GotFocus, txtComments.GotFocus, cmdOrderSearch.GotFocus, _
                    cmdApplyToSearch.GotFocus, cmdCustomerSearch.GotFocus, CmdShipToSearch.GotFocus, _
                    cmdSalespersonSearch.GotFocus, cmdProjectSearch.GotFocus, cmdInstructionsSearch.GotFocus, _
                    cmdLocationSearch.GotFocus, cmdShipViaSearch.GotFocus, cmdTermsSearch.GotFocus, _
                    cmdProfitCenterSearch.GotFocus, cmdCostCenterSearch.GotFocus, cmdCommentsSearch.GotFocus

        Dim lTabIndex As Integer

        If TypeOf sender Is TextBox Then

            Dim txtElement As New TextBox
            txtElement = DirectCast(sender, TextBox)
            lTabIndex = txtElement.TabIndex

        ElseIf TypeOf sender Is ComboBox Then

            Dim txtElement As New ComboBox
            txtElement = DirectCast(sender, ComboBox)
            lTabIndex = txtElement.TabIndex

        ElseIf TypeOf sender Is Button Then

            Dim txtElement As New Button
            txtElement = DirectCast(sender, Button)
            lTabIndex = txtElement.TabIndex

        End If

        blnForceFocus = ForceFocus(lTabIndex)

    End Sub

    ' Trigger search with F7
    Private Sub Element_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
                txtApplyTo.KeyDown, txtCustomer.KeyDown, txtShipTo.KeyDown, _
                cboShipTo.KeyDown, txtSalesPerson.KeyDown, txtProject.KeyDown, _
                txtInstructions.KeyDown, txtLocation.KeyDown, txtComments.KeyDown, _
                cboShipVia.KeyDown, txtTerms.KeyDown, txtProfitCenter.KeyDown, _
                txtCostCenter.KeyDown, txtBillToCountry.KeyDown, txtShipToCountry.KeyDown

        Debug.Print(e.KeyCode)
        Debug.Print(e.KeyData)

        If e.KeyCode = Keys.F7 Then ' e.Control And e.KeyValue = Keys.F Then
            Dim oButton As Button
            oButton = GetSearchControlByControlName(sender.name)
            oButton.PerformClick()
        End If

    End Sub

    ' Sets field color to white or red when data not found in DB
    Private Sub Element_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
                    txtApplyTo.LostFocus, txtCustomer.LostFocus, txtShipTo.LostFocus, _
                    cboShipTo.LostFocus, txtSalesPerson.LostFocus, txtProject.LostFocus, _
                    cboShipVia.LostFocus, txtTerms.LostFocus, txtProfitCenter.LostFocus, _
                    txtCostCenter.LostFocus, txtBillToCountry.LostFocus, txtShipToCountry.LostFocus, _
                    txtLocation.LostFocus 'txtInstructions.LostFocus, txtComments.LostFocus, _

        ' Les instructions et les commentaires ne generent pas d'erreurs

        Dim txtElement As New Object

        If TypeOf sender Is TextBox Then

            txtElement = New TextBox
            txtElement = DirectCast(sender, TextBox)

        ElseIf TypeOf sender Is ComboBox Then

            txtElement = New ComboBox
            txtElement = DirectCast(sender, ComboBox)

        End If

        If ttToolTip.GetToolTip(txtElement) = "Record not on file." Then
            txtElement.BackColor = Color.FromArgb(255, 192, 192)
        Else
            txtElement.BackColor = Color.White
        End If

    End Sub

    ' Autocomplete fields when needed, change fore colors and fill complementary labels
    Private Sub Element_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
                    txtApplyTo.TextChanged, txtCustomer.TextChanged, txtShipTo.TextChanged, _
                    cboShipTo.TextChanged, txtSalesPerson.TextChanged, txtProject.TextChanged, _
                    txtInstructions.TextChanged, txtLocation.TextChanged, txtComments.TextChanged, _
                    cboShipVia.TextChanged, txtTerms.TextChanged, txtProfitCenter.TextChanged, _
                    txtCostCenter.TextChanged, txtBillToCountry.TextChanged, txtShipToCountry.TextChanged

        Dim lTabIndex As Integer

        Dim txtElement As New Object
        Dim lblDescriptionLabel As New Label

        If TypeOf sender Is TextBox Then
            'Dim a1 As TextBox
            txtElement = New TextBox
            txtElement = DirectCast(sender, TextBox)
            lTabIndex = txtElement.TabIndex
            'Attempted to read or write protected memory. This is often an indication that other memory is corrupt.

            Dim txtAutoComplete As New TextBox
            txtAutoComplete = DirectCast(sender, TextBox)

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

        Dim oSearch As New cSearch()
        oSearch.SearchElement = GetSearchElementByControlName(txtElement.name)
        '' oSearch.SetDefaultSearchElement(GetSearchFieldByElement(oSearch.SearchElement), RTrim(txtElement.Text), GetSearchTypeByControlName(txtElement))
        oSearch.SetDefaultSearchElement(GetSearchFieldByElement(oSearch.SearchElement), RTrim(txtElement.Text), "129")
        If Trim(RTrim(txtElement.text)).Length > 2 Then
            oSearch.GetResults(True)
        Else
            oSearch.GetResults()
        End If

        If oSearch.Form.fgSearch.Rows > 2 And Trim(RTrim(sender.text)) <> Trim(RTrim(oSearch.Form.FoundElementValue)) Then
            If TypeOf sender Is TextBox Then
                Dim acscAutoComplete As New AutoCompleteStringCollection()
                For lPos As Integer = 1 To oSearch.Form.fgSearch.Rows - 1
                    acscAutoComplete.Add(oSearch.Form.TextMatrix(lPos, 2))
                Next lPos

                SyncLock txtElement.AutoCompleteCustomSource.SyncRoot
                    txtElement.AutoCompleteCustomSource = acscAutoComplete
                End SyncLock
            Else
                'ttToolTip.SetToolTip(txtElement, "Record not on file.")
            End If
            ttToolTip.SetToolTip(txtElement, "Record not on file.")
        ElseIf oSearch.Form.FoundRow = 1 And Trim(RTrim(sender.text)) = Trim(RTrim(oSearch.Form.FoundElementValue)) Then ' oSearch.Form.fgSearch.Rows = 2 And oSearch.Form.FoundRow = 1 Then
            ' Be careful to always check if the value is the same, else we will enter in a deadlock
            'If oSearch.Form.FoundElementValue <> txtElement.text Then
            '    txtElement.text = oSearch.Form.FoundElementValue
            'End If
            ttToolTip.SetToolTip(txtElement, oSearch.Form.TextMatrix(oSearch.Form.FoundRow, 3))
            If Not (lblDescriptionLabel Is Nothing) Then
                lblDescriptionLabel.Text = oSearch.Form.TextMatrix(oSearch.Form.FoundRow, 3)
            End If
            txtElement.backcolor = Color.White
        Else
            ttToolTip.SetToolTip(txtElement, "Record not on file.")
            'txtElement.backcolor = Color.FromArgb(255, 192, 192)
            If Not (lblDescriptionLabel Is Nothing) Then
                lblDescriptionLabel.Text = ""
            End If
            Dim acscAutoComplete As New AutoCompleteStringCollection()
            SyncLock txtElement.AutoCompleteCustomSource.SyncRoot
                txtElement.AutoCompleteCustomSource = acscAutoComplete
            End SyncLock
        End If

        oSearch = Nothing

        If blnSetTags Then txtElement.Tag = txtElement.Text

        txtElement.forecolor = IIf(txtElement.tag <> txtElement.text, Color.Blue, Color.Black)

    End Sub

    ' Change fore colors on fields with no search
    Private Sub Element_TextChangedNoSearch(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        txtOrder.TextChanged, txtOrderDate.TextChanged, txtShipDate.TextChanged, txtPONumber.TextChanged, _
        txtDiscPct.TextChanged, cboBillToName.TextChanged, txtBillToAddress1.TextChanged, _
        txtBillToAddress2.TextChanged, txtBillToAddress3.TextChanged, txtBillToAddress4.TextChanged, _
        cboShipToName.TextChanged, txtShipToAddress1.TextChanged, txtShipToAddress2.TextChanged, _
        txtShipToAddress3.TextChanged, txtShipToAddress4.TextChanged

        Dim txtElement As New Object

        If TypeOf sender Is TextBox Then
            txtElement = New TextBox
            txtElement = DirectCast(sender, TextBox)

            Dim txtAutoComplete As New TextBox
            txtAutoComplete = DirectCast(sender, TextBox)

        ElseIf TypeOf sender Is ComboBox Then

            txtElement = New ComboBox
            txtElement = DirectCast(sender, ComboBox)

        End If

        If blnSetTags Then txtElement.Tag = txtElement.Text

        txtElement.forecolor = IIf(txtElement.tag <> txtElement.text, Color.Blue, Color.Black)

    End Sub

    Private Sub cboCustomerNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCustomerNo.GotFocus

        If blnForceFocus Then Exit Sub

        If cboCustomerNo.Items.Count = 0 Then
            Dim oSearch As New cSearch("AR-CUSTOMER", False, True)
            For lPos As Integer = 0 To oSearch.Form.TextMatrix.Rows - 1
                cboCustomerNo.Items.Add(oSearch.Form.TextMatrix(lPos, 2))
            Next
            oSearch = Nothing
        End If

    End Sub

    ' fills the combobox with the adresses found
    Private Sub cboShipTo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboShipTo.GotFocus

        If blnForceFocus Then Exit Sub

        Dim osearch As New cSearch()
        osearch.SearchElement = "AR-ALT-ADDR-COD"
        osearch.SetDefaultSearchElement("RTRIM(ARALTADR_SQL.cus_no)", RTrim(txtCustomer.Text), "200")
        osearch.Search()
        cboShipTo.Items.Clear()
        cboShipTo.Items.Add(" ")
        For lPos As Integer = 1 To osearch.Form.TextMatrix.Rows - 1
            cboShipTo.Items.Add(osearch.Form.TextMatrix(lPos, 2))
        Next
        osearch = Nothing

    End Sub

    ' Combobox for ship to, we force the customer no to the search filter.
    Private Sub CmdShipToSearch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdShipToSearch.Click

        If blnForceFocus Then Exit Sub

        Dim oSearch As New cSearch()
        oSearch.SearchElement = "AR-ALT-ADDR-COD"
        oSearch.SetDefaultSearchElement("RTRIM(ARALTADR_SQL.cus_no)", RTrim(txtCustomer.Text), "200")
        oSearch.Form.ShowDialog()
        If oSearch.Form.FoundElementValue <> String.Empty Then
            'Dim oSearch As New cSearch("AR-ALT-ADDR-COD")
            txtShipTo.Text = oSearch.Form.FoundElementValue
            cboShipTo.Text = txtShipTo.Text
        End If
        oSearch = Nothing

    End Sub

    ' Puts the selected date from the calendar info the linked textbox.
    Private Sub mcCalendar_DateSelected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles mcCalendar.DateSelected

        Select Case m_DateButton.Name
            Case "cmdDateTP"
                txtOrderDate.Text = mcCalendar.SelectionRange.Start

            Case "cmdShipDateTP"
                txtShipDate.Text = mcCalendar.SelectionRange.Start

        End Select
        mcCalendar.Visible = False
        m_DateButton = Nothing

    End Sub

    ' Checks if a customer is valid, if so then sets default values for it.
    Private Sub txtCustomer_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomer.Validated

        If txtCustomer.Text <> txtCustomer.Tag Then
            If sender.backcolor = Color.White Then
                Call SetDefaultCustomerValues()
            End If

        End If
        txtCustomer.Tag = txtCustomer.Text

    End Sub

    ' Sets the order number white or pink, is a required field.
    Private Sub txtOrder_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOrder.TextChanged

        If String.IsNullOrEmpty(e.ToString) Then
            txtOrder.BackColor = Color.FromArgb(255, 192, 192)
        Else
            txtOrder.BackColor = Color.White
        End If

    End Sub

    ' This proc checks the tabindex of the caller and compare it to needed fields.
    ' Also, completes some fields depending on position from tabindex order.
    Private Function ForceFocus(ByVal piTabIndex As Integer) As Boolean

        ForceFocus = False

        ' Set focus to Order number if not entered
        If piTabIndex >= txtOrderDate.TabIndex And (txtOrder.Text = "" Or txtOrder.BackColor = Color.FromArgb(255, 192, 192)) Then
            txtOrder.Focus()
            txtOrder.BackColor = Color.FromArgb(255, 192, 192)
            ttToolTip.SetToolTip(txtOrder, "Data must be entered here.")
            ForceFocus = True

            Exit Function

        End If

        ' if no date has been entered and click after, then enter today as date
        If piTabIndex >= txtApplyTo.TabIndex And (txtOrderDate.Text = "") Then
            txtOrderDate.Text = Date.Now.ToShortDateString
        End If

        ' if no ship date has been entered and click after, then enter today as date
        If piTabIndex >= txtCustomer.TabIndex And (txtShipDate.Text = "") Then
            txtShipDate.Text = txtOrderDate.Text ' Date.Now.ToShortDateString
        End If

        ' If no customer has been selected, then you can't select a ship to address.
        If piTabIndex >= txtShipTo.TabIndex And piTabIndex <= cmdShipToNotes.TabIndex And (txtCustomer.Text = "" Or txtCustomer.BackColor = Color.FromArgb(255, 192, 192)) Then

            txtCustomer.Focus()
            txtCustomer.BackColor = Color.FromArgb(255, 192, 192)
            ttToolTip.SetToolTip(txtCustomer, "Data must be entered here.")
            ForceFocus = True

            Exit Function

        End If

        If piTabIndex >= txtPONumber.TabIndex Then

            ' Si le ship to est vide, on entre celui par defaut.
            If Not (txtCustomer.Text = String.Empty Or txtCustomer.BackColor <> Color.White) Then

                If cboShipToName.Text = String.Empty Then

                    blnSetTags = True

                    cboShipToName.Text = cboBillToName.Text
                    txtShipToAddress1.Text = txtBillToAddress1.Text
                    txtShipToAddress2.Text = txtBillToAddress2.Text
                    txtShipToAddress3.Text = txtBillToAddress3.Text
                    txtShipToAddress4.Text = txtBillToAddress4.Text
                    txtShipToCountry.Text = txtBillToCountry.Text
                    cboShipTo.Text = "SAME"

                    blnSetTags = False

                End If

            End If

        End If

    End Function

    ' Returns the label to show complimentary data on.
    Public Function GetDescriptionLabelByElement(ByVal txtElementName As String) As Label

        GetDescriptionLabelByElement = Nothing

        Select Case txtElementName

            'Case "txtOrder"
            '    GetSearchElementByControlName = "OE-ORDER-NO"

            Case "txtLocation"
                GetDescriptionLabelByElement = lblLocationDesc

            Case "cboShipVia"
                GetDescriptionLabelByElement = lblShipViaDesc

            Case "txtTerms"
                GetDescriptionLabelByElement = lblTermsDesc

            Case "txtBillToCountry"
                GetDescriptionLabelByElement = lblBillToCountryDesc

            Case "txtShipToCountry"
                GetDescriptionLabelByElement = lblShipToCountryDesc

        End Select

    End Function

    ' Returns the textbox or the combobox to apply the search from the search button
    Public Function GetElementByControlName(ByVal txtElement As String) As Object

        GetElementByControlName = ""

        Select Case txtElement

            Case "cmdOrderSearch" ' "txtOrder"
                GetElementByControlName = txtOrder

            Case "txtApplyTo", "cmdApplyToSearch"
                GetElementByControlName = txtApplyTo

            Case "txtCustomer", "cmdCustomerSearch"
                GetElementByControlName = txtCustomer

            Case "txtSalesPerson", "cmdSalespersonSearch"
                GetElementByControlName = txtSalesPerson

            Case "txtProject", "cmdProjectSearch"
                GetElementByControlName = txtProject

            Case "txtLocation", "cmdLocationSearch"
                GetElementByControlName = txtLocation

            Case "txtTerms", "cmdTermsSearch"
                GetElementByControlName = txtTerms

            Case "txtProfitCenter", "cmdProfitCenterSearch"
                GetElementByControlName = txtProfitCenter

            Case "txtCostCenter", "cmdCostCenterSearch"
                GetElementByControlName = txtCostCenter

            Case "txtInstructions", "cmdInstructionsSearch"
                GetElementByControlName = txtInstructions

            Case "txtComments", "cmdCommentsSearch"
                GetElementByControlName = txtComments

            Case "txtBillToCountry", "cmdBillToCountrySearch"
                GetElementByControlName = txtBillToCountry

            Case "txtShipToCountry", "cmdShipToCountrySearch"
                GetElementByControlName = txtShipToCountry

            Case "cboShipTo"
                GetElementByControlName = cboShipTo

            Case "cboShipVia", "cmdShipViaSearch"
                GetElementByControlName = cboShipVia

        End Select

    End Function

    ' Get search field from SEARCHES_SQL for a control name (to link with FieldID)
    Public Function GetSearchFieldByElement(ByVal pstrSearchElement As String) As String

        GetSearchFieldByElement = ""

        Select Case pstrSearchElement

            Case "AR-ALT-ADDR-COD"
                GetSearchFieldByElement = "ARALTADR_SQL.cus_alt_adr_cd"

            Case "AR-CUSTOMER"
                GetSearchFieldByElement = "arcusfil_sql.cus_no"

            Case "AR-SALESPERSON"
                GetSearchFieldByElement = "arslmfil_sql.humres_id"

            Case "AR-SHIP-VIA"
                GetSearchFieldByElement = "SYCDEFIL_SQL.sy_code"

            Case "AR-TERMS-CODE"
                GetSearchFieldByElement = "SYTRMFIL_SQL.term_code"

            Case "IM-LANDCODE"
                GetSearchFieldByElement = "land.landcode"

            Case "IM-LOCATION"
                GetSearchFieldByElement = "IMLOCFIL_SQL.loc"

            Case "OE-COMMENT-CD-2"
                GetSearchFieldByElement = "OECMTCDE_SQL.cmt_cd"

            Case "OE-HDR-HST-INV"
                GetSearchFieldByElement = "OEHDRHST_SQL.inv_no"

            Case "OE-ORDER-NO"
                GetSearchFieldByElement = "OEORDHDR_SQL.cus_no"

            Case "SY-ACCOUNT2"
                GetSearchFieldByElement = "kstdr.kstdrcode"

            Case "SY-ACCOUNT3"
                GetSearchFieldByElement = "kstpl.kstplcode"

            Case "SY-JOB-CODE"
                GetSearchFieldByElement = "jobfile_sql.job_no"

        End Select

    End Function

    ' Returns the search button control associated with a textbox or a combobox
    Public Function GetSearchControlByControlName(ByVal txtElement As String) As Button

        GetSearchControlByControlName = New Button

        Select Case txtElement

            'Case "txtOrder" 
            '    GetSearchControlByControlName = cmdOrderSearch

            Case "txtApplyTo"
                GetSearchControlByControlName = cmdApplyToSearch

            Case "txtCustomer"
                GetSearchControlByControlName = cmdCustomerSearch

            Case "txtSalesPerson"
                GetSearchControlByControlName = cmdSalespersonSearch

            Case "txtProject"
                GetSearchControlByControlName = cmdProjectSearch

            Case "txtLocation"
                GetSearchControlByControlName = cmdLocationSearch

            Case "txtTerms"
                GetSearchControlByControlName = cmdTermsSearch

            Case "txtProfitCenter"
                GetSearchControlByControlName = cmdProfitCenterSearch

            Case "txtCostCenter"
                GetSearchControlByControlName = cmdCostCenterSearch

            Case "txtInstructions"
                GetSearchControlByControlName = cmdInstructionsSearch

            Case "txtComments"
                GetSearchControlByControlName = cmdCommentsSearch

            Case "txtBillToCountry"
                GetSearchControlByControlName = cmdBillToCountrySearch

            Case "txtShipToCountry"
                GetSearchControlByControlName = cmdShipToCountrySearch

            Case "cboShipTo"
                GetSearchControlByControlName = CmdShipToSearch

            Case "cboShipVia"
                GetSearchControlByControlName = cmdShipViaSearch

        End Select

    End Function

    ' Get Field_ID for SEARCHES_SQL table for a control name
    Public Function GetSearchElementByControlName(ByVal txtElement As String) As String

        GetSearchElementByControlName = ""

        Select Case txtElement

            Case "cmdOrderSearch" ' "txtOrder"
                GetSearchElementByControlName = "OE-ORDER-NO"

            Case "txtApplyTo", "cmdApplyToSearch"
                GetSearchElementByControlName = "OE-HDR-HST-INV"

            Case "txtCustomer", "cmdCustomerSearch"
                GetSearchElementByControlName = "AR-CUSTOMER"

            Case "txtSalesPerson", "cmdSalespersonSearch"
                GetSearchElementByControlName = "AR-SALESPERSON"

            Case "txtProject", "cmdProjectSearch"
                GetSearchElementByControlName = "SY-JOB-CODE"

            Case "txtLocation", "cmdLocationSearch"
                GetSearchElementByControlName = "IM-LOCATION"

            Case "txtTerms", "cmdTermsSearch"
                GetSearchElementByControlName = "AR-TERMS-CODE"

            Case "txtProfitCenter", "cmdProfitCenterSearch"
                GetSearchElementByControlName = "SY-ACCOUNT2"

            Case "txtCostCenter", "cmdCostCenterSearch"
                GetSearchElementByControlName = "SY-ACCOUNT3"

            Case "txtInstructions", "cmdInstructionsSearch"
                GetSearchElementByControlName = "OE-COMMENT-CD-2"

            Case "txtComments", "cmdCommentsSearch"
                GetSearchElementByControlName = "OE-COMMENT-CD-2"

            Case "txtBillToCountry", "cmdBillToCountrySearch"
                GetSearchElementByControlName = "IM-LANDCODE"

            Case "txtShipToCountry", "cmdShipToCountrySearch"
                GetSearchElementByControlName = "IM-LANDCODE"

            Case "cboShipTo"
                GetSearchElementByControlName = "AR-ALT-ADDR-COD"

            Case "cboShipVia", "cmdShipViaSearch"
                GetSearchElementByControlName = "AR-SHIP-VIA"

        End Select

    End Function

    ' Reset every field after Customer name to String.Empty
    Private Sub ResetCustomerFields()

        txtPONumber.Text = String.Empty
        txtSalesPerson.Text = String.Empty
        txtProject.Text = String.Empty

        cboShipTo.Text = String.Empty
        txtLocation.Text = String.Empty
        cboShipVia.Text = String.Empty
        txtTerms.Text = String.Empty
        txtDiscPct.Text = String.Empty
        txtProfitCenter.Text = String.Empty
        txtCostCenter.Text = String.Empty
        txtInstructions.Text = String.Empty
        txtComments.Text = String.Empty

        cboBillToName.Text = String.Empty
        txtBillToAddress1.Text = String.Empty
        txtBillToAddress2.Text = String.Empty
        txtBillToAddress3.Text = String.Empty
        txtBillToAddress4.Text = String.Empty
        txtBillToCountry.Text = String.Empty
        lblBillToCountryDesc.Text = String.Empty

        cboShipToName.Text = String.Empty
        txtShipToAddress1.Text = String.Empty
        txtShipToAddress2.Text = String.Empty
        txtShipToAddress3.Text = String.Empty
        txtShipToAddress4.Text = String.Empty
        txtShipToCountry.Text = String.Empty
        lblShipToCountryDesc.Text = String.Empty

    End Sub

    ' Set default values for customer from ARCUSFIL_SQL record for every field after customer name
    Private Sub SetDefaultCustomerValues()

        ' Must first validate if items lines are included, then
        ' must ask "will delete the items selected. will you continue'
        ' because the items are company specific, we must do it that way
        ' and by no mean do a bypass.

        ' Will set the tags to the value entered, 
        ' every change afterwards will turn the field to blue if different as in Macola.
        blnSetTags = True

        Dim strSql As String
        strSql = "" & _
        "SELECT     TOP 1 * " & _
        "FROM       ARCUSFIL_SQL WITH (Nolock) " & _
        "WHERE      RTRIM(ISNULL(CUS_NO, '')) = '" & RTrim(txtCustomer.Text) & "' " & _
        "ORDER BY   CUS_NO ASC "

        ' Verifier ceux qui ont des Cus_Alt_Adr_Cd

        Dim rsCustomer As New ADODB.Recordset
        rsCustomer.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCustomer.RecordCount <> 0 Then

            cboShipTo.Text = String.Empty

            txtPONumber.Text = String.Empty
            txtSalesPerson.Text = IIf(rsCustomer.Fields("Slspsn_No").Equals(System.DBNull.Value), "", rsCustomer.Fields("Slspsn_No").Value)
            txtProject.Text = String.Empty

            txtLocation.Text = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Loc").Value), "", rsCustomer.Fields("Loc").Value)
            cboShipVia.Text = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Ship_Via_Cd").Value), "", rsCustomer.Fields("Ship_Via_Cd").Value)
            txtTerms.Text = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Ar_Terms_Cd").Value), "", rsCustomer.Fields("Ar_Terms_Cd").Value)
            txtDiscPct.Text = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Dsc_Pct").Value), "", rsCustomer.Fields("Dsc_Pct").Value)
            txtProfitCenter.Text = String.Empty
            txtCostCenter.Text = String.Empty

            txtInstructions.Text = String.Empty
            txtComments.Text = String.Empty

            cboBillToName.Text = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Cus_Name").Value), "", rsCustomer.Fields("Cus_Name").Value)
            txtBillToAddress1.Text = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Addr_1").Value), "", rsCustomer.Fields("Addr_1").Value)
            txtBillToAddress2.Text = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Addr_2").Value), "", rsCustomer.Fields("Addr_2").Value)
            txtBillToAddress3.Text = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Addr_3").Value), "", rsCustomer.Fields("Addr_3").Value)
            txtBillToAddress4.Text = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("City").Value), "", rsCustomer.Fields("City").Value) & ", " & _
                IIf(System.DBNull.Value.Equals(rsCustomer.Fields("State").Value), "", rsCustomer.Fields("State").Value) & " " & _
                IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Zip").Value), "", rsCustomer.Fields("Zip").Value)
            txtBillToCountry.Text = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Country").Value), "", rsCustomer.Fields("Country").Value)
            'txtBillToCountryDesc.Text = rsCustomer.Fields("").Value 'will populate on country change

            cboShipToName.Text = String.Empty
            txtShipToAddress1.Text = String.Empty
            txtShipToAddress2.Text = String.Empty
            txtShipToAddress3.Text = String.Empty
            txtShipToAddress4.Text = String.Empty
            txtShipToCountry.Text = String.Empty
            lblShipToCountryDesc.Text = String.Empty
        Else
            Call ResetCustomerFields()
        End If

        rsCustomer.Close()

        blnSetTags = False

    End Sub

End Class
