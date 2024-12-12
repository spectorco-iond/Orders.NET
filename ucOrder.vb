Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Imports Microsoft.VisualBasic.PowerPacks
Imports System.IO
Imports System.Data.SqlClient
Imports System.Threading

Friend Class ucOrder
    Inherits System.Windows.Forms.UserControl

    Dim mailer As Object

    Dim ofrmOEFlash As frmOEFlash
    Dim ofrmCustomerConfig As frmCustomerConfig

    Public Event White_Flag_Customer(ByVal sender As Object, ByVal e As System.EventArgs)

    ' VARIABLES AND ENUM DEFINITIONS xxxxxxxxxxxxxxxxxxxxxxxxxxxxx

#Region "Private variables ################################################"
    'Private m_oOrder As New cOrder
    Private blnLoading As Boolean = False

    Private _Item_No As String
    Private blnInvalidate As Boolean = False
    Private iRowPos As Integer = 0
    Private iColPos As Integer = 0

    Private blnForceFocus As Boolean = False
    Private m_DateButton As Button

    Private qtyEdi As Integer = 0

    'Private dsDataSet As New DataSet()

    Private blnSetTags As Boolean = False

    Private acscAutoComplete As New AutoCompleteStringCollection()

    Delegate Sub myDelegate() ' ByVal pintPanel As Integer, ByVal pstrText As String)
    Delegate Sub LoadOrderFromOEIDelegate(ByVal pstrOrd_No As String)

    Private WithEvents TestWorker As System.ComponentModel.BackgroundWorker

    Private m_sbptPanel As Integer
    Private m_sbptText As String
    Dim backgroundThread As Thread
    Dim IdUser, xml_Access As String

#End Region


#Region "Private enum #####################################################"

    ' Order line column positions
    Private Enum _Columns
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
        _Description ' 10
        _Description2
        _UOM
        _TaxSchd
        _ECMRevision
        _CreateModifyKit
        _NoteText
        _BackorderItem
        _Selected
        _Recalc
        _LineComments ' 20
        _Route
    End Enum

#End Region


    ' CODE xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

#Region "Private elements control functions ###############################"

    ' Returns the label to show complimentary data on.
    Private Function GetDescriptionLabelByElement(ByVal txtElementName As String) As Label

        GetDescriptionLabelByElement = Nothing

        Try
            Select Case txtElementName

                'Case "txtOrder"
                '    GetSearchElementByControlName = "OE-ORDER-NO"

                Case "txtMfg_Loc"
                    GetDescriptionLabelByElement = lblLocationDesc

                Case "txtShip_Via_Cd"
                    GetDescriptionLabelByElement = lblShipViaDesc

                Case "txtAr_Terms_Cd"
                    GetDescriptionLabelByElement = lblTermsDesc

                Case "txtBill_To_Country"
                    GetDescriptionLabelByElement = lblBillToCountryDesc

                Case "txtShip_To_Country"
                    GetDescriptionLabelByElement = lblShipToCountryDesc

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

                Case "cmdOrderSearch", "cmdOrd_No_Notes"
                    GetElementByControlName = txtOrd_No

                    'Case "txtApplyTo", "cmdApplyToSearch"
                    '    GetElementByControlName = txtApplyTo

                Case "txtCus_No", "cmdCustomerSearch", "cmdCus_No_Notes"
                    GetElementByControlName = txtCus_No

                Case "txtSlspsn_No", "cmdSalespersonSearch", "cmdSlspsn_No_Notes"
                    GetElementByControlName = txtSlspsn_No

                Case "txtShip_Link", "cmdShip_LinkSearch", "cmdShip_LinkEdit"
                    GetElementByControlName = txtShip_Link

                Case "txtJob_No", "cmdProjectSearch", "cmdJob_No_Notes"
                    GetElementByControlName = txtJob_No

                Case "txtMfg_Loc", "cmdLocationSearch", "cmdMfg_Loc_Notes"
                    GetElementByControlName = txtMfg_Loc

                Case "txtAr_Terms_Cd", "cmdTermsSearch", "cmdAr_Terms_Cd_Notes"
                    GetElementByControlName = txtAr_Terms_Cd

                Case "txtProfit_Center", "cmdProfitCenterSearch", "cmdProfit_Center_Notes"
                    GetElementByControlName = txtProfit_Center

                Case "txtDept", "cmdCostCenterSearch", "cmdDept_Notes"
                    GetElementByControlName = txtDept

                Case "txtShip_Instruction_1", "cmdInstructionsSearch"
                    GetElementByControlName = txtShip_Instruction_1

                Case "txtShip_Instruction_2", "cmdCommentsSearch"
                    GetElementByControlName = txtShip_Instruction_2

                Case "txtBill_To_Country", "cmdBillToCountrySearch", "cmdBill_To_Country_Notes"
                    GetElementByControlName = txtBill_To_Country

                Case "txtShip_To_Country", "cmdShipToCountrySearch", "cmdShip_To_Country_Notes"
                    GetElementByControlName = txtShip_To_Country

                Case "txtCus_Alt_Adr_Cd", "cmdCus_Alt_Adr_Cd_Notes"
                    GetElementByControlName = txtCus_Alt_Adr_Cd

                Case "txtShip_Via_Cd", "cmdShipViaSearch", "cmdShip_Via_Cd_Notes"
                    GetElementByControlName = txtShip_Via_Cd

                Case "txtEnd_User"
                    GetElementByControlName = txtEnd_User

                Case "txtUser_Def_Fld_1"
                    GetElementByControlName = txtShip_Via_Cd

                Case "txtUser_Def_Fld_3", "cmdUser_Def_Fld_3Search"
                    GetElementByControlName = txtUser_Def_Fld_3

                Case "txtUser_Def_Fld_4", "cmdUser_Def_Fld_4Search"
                    GetElementByControlName = txtUser_Def_Fld_4

                Case "txtUser_Def_Fld_5", "cmdUser_Def_Fld_5Search"
                    GetElementByControlName = txtUser_Def_Fld_5

                Case "txtShip_To_Name", "cmdShip_To_NameSearch"
                    GetElementByControlName = txtShip_To_Name

                Case "txtProg_Spector_Cd", "cmdProg_Spector_CdSearch"
                    GetElementByControlName = txtProg_Spector_Cd

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

#End Region


#Region "Private DateTimePicker control functions #########################"

    ' Date buttons click - open DateControl for element
    Private Sub Element_DateButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDateTP.Click, cmdShipDateTP.Click, cmdInHandDate.Click


        If blnForceFocus Then Exit Sub

        Try
            m_DateButton = New Button
            m_DateButton = DirectCast(sender, Button)

            mcCalendar.Top = m_DateButton.Top
            mcCalendar.Left = m_DateButton.Left
            mcCalendar.Visible = True
            mcCalendar.SetDate(Date.Now)

            Select Case m_DateButton.Name
                Case "cmdDateTP"
                    If IsDate(txtOrd_Dt.Text) Then mcCalendar.SetDate(txtOrd_Dt.Text)

                Case "cmdShipDateTP"
                    If IsDate(txtOrd_Dt_Shipped.Text) Then mcCalendar.SetDate(txtOrd_Dt_Shipped.Text)

            End Select
            mcCalendar.Select()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Returns the search button control associated with a textbox or a combobox
    Private Function GetDateControlByControlName(ByVal txtElement As String) As Button

        GetDateControlByControlName = New Button

        Try
            Select Case txtElement

                Case "txtOrd_Dt"
                    GetDateControlByControlName = cmdDateTP

                Case "txtOrd_Dt_Shipped"
                    GetDateControlByControlName = cmdShipDateTP

                Case "txtInHandsDate"
                    GetDateControlByControlName = cmdInHandDate

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Puts the selected date from the calendar info the linked textbox.
    Private Sub mcCalendar_DateSelected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles mcCalendar.DateSelected


        Try
            'If mcCalendar.SelectionRange.Start.Year < (Now.Year - 1) Then
            '    Throw New OEException(OEError.Year_Cannot_Be_LT_Previous_Year)
            'End If

            'If mcCalendar.SelectionRange.Start.Year > (Now.Year + 1) Then
            '    Throw New OEException(OEError.Year_Cannot_Be_GT_Follow_Year)
            'End If

            Select Case m_DateButton.Name
                Case "cmdDateTP"
                    txtOrd_Dt.Text = mcCalendar.SelectionRange.Start
                    txtOrd_Dt.Focus()

                Case "cmdShipDateTP"
                    txtOrd_Dt_Shipped.Text = mcCalendar.SelectionRange.Start
                    txtOrd_Dt_Shipped.Focus()
                Case "cmdInHandDate"
                    txtInHandsDate.Text = mcCalendar.SelectionRange.Start
                    txtInHandsDate.Focus()
            End Select

        Catch oe_er As OEException
            MsgBox(oe_er.Message)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        Finally
            mcCalendar.Visible = False
            m_DateButton = Nothing
        End Try

    End Sub

    ' When date time picker gets Escape button, it closes.
    Private Sub mcCalendar_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mcCalendar.KeyDown

        Try
            If e.KeyCode = Keys.Escape Then ' e.Control And e.KeyValue = Keys.F Then
                mcCalendar.Visible = False
            End If

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

                Case "cmdOrd_No_Notes" ' "txtOrder"
                    GetNotesElementByControlName = "OE-ORDER-NO"

                Case "cmdCus_No_Notes"
                    GetNotesElementByControlName = "AR-CUSTOMER"

                Case "cmdSlspsn_No_Notes"
                    GetNotesElementByControlName = "AR-SALESPERSON"

                Case "cmdJob_No_Notes"
                    GetNotesElementByControlName = "SY-JOB-CODE"

                Case "cmdMfg_Loc_Notes"
                    GetNotesElementByControlName = "IM-LOCATION"

                Case "cmdAr_Terms_Cd_Notes"
                    GetNotesElementByControlName = "AR-TERMS-CODE"

                Case "cmdProfit_Center_Notes"
                    GetNotesElementByControlName = "SY-ACCOUNT2"

                Case "cmdDept_Notes"
                    GetNotesElementByControlName = "SY-ACCOUNT3"

                Case "cmdBill_To_Country_Notes", "cmdShip_To_Country_Notes"
                    GetNotesElementByControlName = "OE-COMMENT-CD-2"

                Case "cmdCus_Alt_Adr_Cd_Notes"
                    GetNotesElementByControlName = "AR-ALT-ADDR-COD"

                Case "cmdShip_Via_Cd_Notes"
                    GetNotesElementByControlName = "AR-SHIP-VIA"

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

                Case "txtOrd_No"
                    GetNotesControlByControlName = cmdOrd_No_Notes

                Case "txtCus_No"
                    GetNotesControlByControlName = cmdCus_No_Notes

                Case "txtSlspsn_No"
                    GetNotesControlByControlName = cmdSlspsn_No_Notes

                Case "txtJob_No"
                    GetNotesControlByControlName = cmdJob_No_Notes

                Case "txtMfg_Loc"
                    GetNotesControlByControlName = cmdMfg_Loc_Notes

                Case "txtAr_Terms_Cd"
                    GetNotesControlByControlName = cmdAr_Terms_Cd_Notes

                Case "txtProfit_Center"
                    GetNotesControlByControlName = cmdProfit_Center_Notes

                Case "txtDept"
                    GetNotesControlByControlName = cmdDept_Notes

                Case "txtBill_To_Country"
                    GetNotesControlByControlName = cmdBill_To_Country_Notes

                Case "txtShip_To_Country"
                    GetNotesControlByControlName = cmdShip_To_Country_Notes

                Case "txtCus_Alt_Adr_Cd"
                    GetNotesControlByControlName = cmdCus_Alt_Adr_Cd_Notes

                Case "txtShip_Via_Cd"
                    GetNotesControlByControlName = cmdShip_Via_Cd_Notes

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Search buttons click - open search window for element
    Private Sub Notes_Element_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _
        cmdOrd_No_Notes.Click, cmdCus_No_Notes.Click, cmdCus_Alt_Adr_Cd_Notes.Click, _
        cmdSlspsn_No_Notes.Click, cmdJob_No_Notes.Click, cmdMfg_Loc_Notes.Click, _
        cmdShip_Via_Cd_Notes.Click, cmdAr_Terms_Cd_Notes.Click, cmdProfit_Center_Notes.Click, _
        cmdDept_Notes.Click, cmdBill_To_Country_Notes.Click, cmdShip_To_Country_Notes.Click

        If blnForceFocus Then Exit Sub

        Dim btnElement As Button

        Try

            If TypeOf eventSender Is Button Then

                btnElement = New Button
                btnElement = DirectCast(eventSender, Button)

                Dim txtElement As Object
                txtElement = GetElementByControlName(btnElement.Name)

                Dim oNotes As New cNotes(GetNotesElementByControlName(btnElement.Name), txtElement.Text.ToString)
                oNotes = Nothing

            End If
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Private Search control functions #################################"

    ' Combobox for ship to, we force the customer no to the search filter.
    Private Sub CmdShipToSearch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdShipToSearch.Click

        If blnForceFocus Then Exit Sub

        Try

            Dim oSearch As New cSearch()
            oSearch.SearchElement = "AR-ALT-ADDR-COD"
            oSearch.SetDefaultSearchElement("RTRIM(ARALTADR_SQL.cus_no)", RTrim(txtCus_No.Text), "200")
            oSearch.Form.ShowDialog()
            If oSearch.Form.FoundElementValue <> String.Empty Then
                'Dim oSearch As New cSearch("AR-ALT-ADDR-COD")
                txtCus_Alt_Adr_Cd.Text = oSearch.Form.FoundElementValue
                'cboShipTo.Text = txtCus_Alt_Adr_Cd.Text
            End If
            oSearch.Dispose()
            oSearch = Nothing
            txtCus_Alt_Adr_Cd.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Combobox for ship to, we force the customer no to the search filter.
    Private Sub CmdShip_To_Name_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShip_To_NameSearch.Click

        If blnForceFocus Then Exit Sub

        Try

            Dim oSearch As New cSearch()
            oSearch.SearchElement = "OEI-SHIP-ADDR"
            oSearch.SetDefaultSearchElement("RTRIM(OEHdrHst_SQL.Cus_No)", RTrim(txtCus_No.Text), "200")
            oSearch.Form.ShowDialog()
            If oSearch.Form.FoundElementValue <> String.Empty Then
                'Dim oSearch As New cSearch("AR-ALT-ADDR-COD")
                Dim lIndex As String = oSearch.Form.FoundElementIndex
                If IsNumeric(lIndex) Then
                    Dim dt As DataTable
                    Dim db As New cDBA
                    Dim strsql As String
                    strsql = _
                    "SELECT     ISNULL(Ship_To_Name, '')    AS Ship_To_Name, " & _
                    "           ISNULL(Ship_To_Addr_1, '')  AS Ship_To_Addr_1, " & _
                    "           ISNULL(Ship_To_Addr_2, '')  AS Ship_To_Addr_2, " & _
                    "           ISNULL(Ship_To_Addr_3, '')  AS Ship_To_Addr_3, " & _
                    "           ISNULL(Ship_To_Addr_4, '')  AS Ship_To_Addr_4, " & _
                    "           ISNULL(Ship_To_City, '')    AS Ship_To_City, " & _
                    "           ISNULL(Ship_To_State, '')   AS Ship_To_State, " & _
                    "           ISNULL(Ship_To_Zip, '')     AS Ship_To_Zip, " & _
                    "           ISNULL(Ship_To_Country, '') AS Ship_To_Country " & _
                    "FROM       OEHdrHst_Sql WITH (Nolock) " & _
                    "WHERE      ID = " & lIndex

                    dt = db.DataTable(strsql)

                    If dt.Rows.Count <> 0 Then

                        txtShip_To_Name.Text = dt.Rows(0).Item("Ship_To_Name").ToString
                        txtShip_To_Addr_1.Text = dt.Rows(0).Item("Ship_To_Addr_1").ToString
                        txtShip_To_Addr_2.Text = dt.Rows(0).Item("Ship_To_Addr_2").ToString
                        txtShip_To_Addr_3.Text = dt.Rows(0).Item("Ship_To_Addr_3").ToString
                        txtShip_To_Addr_4.Text = dt.Rows(0).Item("Ship_To_Addr_4").ToString
                        txtShip_To_City.Text = dt.Rows(0).Item("Ship_To_City").ToString
                        txtShip_To_State.Text = dt.Rows(0).Item("Ship_To_State").ToString
                        txtShip_To_Zip.Text = dt.Rows(0).Item("Ship_To_Zip").ToString
                        txtShip_To_Country.Text = dt.Rows(0).Item("Ship_To_Country").ToString

                    End If

                End If
                'txtShip_To_Name.Text = oSearch.Form.FoundElementValue
                'cboShipTo.Text = txtCus_Alt_Adr_Cd.Text
            End If
            oSearch.Dispose()
            oSearch = Nothing
            txtShip_To_Name.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Combobox for ship to, we force the customer no to the search filter.
    Private Sub CmdShip_Link_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShip_LinkSearch.Click

        If blnForceFocus Then Exit Sub

        Try

            Dim oSearch As New cSearch()
            oSearch.SearchElement = "OEI-SHIP-LINK"
            oSearch.SetDefaultSearchElement("RTRIM(OEOrdHdr_SQL.cus_no)", RTrim(txtCus_No.Text), "200")
            oSearch.Form.ShowDialog()
            If oSearch.Form.FoundElementValue <> String.Empty Then
                'Dim oSearch As New cSearch("AR-ALT-ADDR-COD")
                txtShip_Link.Text = oSearch.Form.FoundElementValue
                'cboShipTo.Text = txtCus_Alt_Adr_Cd.Text
            End If
            oSearch.Dispose()
            oSearch = Nothing
            txtShip_Link.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function AutoCompleteData(ByVal txtElement As String) As Boolean

        AutoCompleteData = False

        Try
            Select Case txtElement

                Case "txtUser_Def_Fld_3"
                    AutoCompleteData = True

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function
    ' Returns the search button control associated with a textbox or a combobox
    Private Function GetSearchControlByControlName(ByVal txtElement As String) As Button

        GetSearchControlByControlName = New Button

        Try
            Select Case txtElement

                'Case "txtOrder" 
                '    GetSearchControlByControlName = cmdOrderSearch

                'Case "txtApplyTo"
                '    GetSearchControlByControlName = cmdApplyToSearch

                Case "txtCus_No"
                    GetSearchControlByControlName = cmdCustomerSearch

                Case "txtSlspsn_No"
                    GetSearchControlByControlName = cmdSalespersonSearch

                Case "txtShip_Link"
                    GetSearchControlByControlName = cmdShip_LinkSearch

                Case "txtJob_No"
                    GetSearchControlByControlName = cmdProjectSearch

                Case "txtMfg_Loc"
                    GetSearchControlByControlName = cmdLocationSearch

                Case "txtAr_Terms_Cd"
                    GetSearchControlByControlName = cmdTermsSearch

                Case "txtProfit_Center"
                    GetSearchControlByControlName = cmdProfitCenterSearch

                Case "txtDept"
                    GetSearchControlByControlName = cmdDeptSearch

                Case "txtShip_Instruction_1"
                    GetSearchControlByControlName = cmdInstructionsSearch

                Case "txtShip_Instruction_2"
                    GetSearchControlByControlName = cmdCommentsSearch

                Case "txtBill_To_Country"
                    GetSearchControlByControlName = cmdBillToCountrySearch

                Case "txtShip_To_Country"
                    GetSearchControlByControlName = cmdShipToCountrySearch

                Case "txtCus_Alt_Adr_Cd"
                    GetSearchControlByControlName = CmdShipToSearch

                Case "txtShip_Via_Cd"
                    GetSearchControlByControlName = cmdShipViaSearch

                Case "txtUser_Def_Fld_3"
                    GetSearchControlByControlName = cmdUser_Def_Fld_3Search

                Case "txtUser_Def_Fld_4"
                    GetSearchControlByControlName = cmdUser_Def_Fld_4Search

                Case "txtUser_Def_Fld_5"
                    GetSearchControlByControlName = cmdUser_Def_Fld_5Search

                Case "txtShip_To_Name"
                    GetSearchControlByControlName = cmdShip_To_NameSearch

                Case "txtProg_Spector_Cd"
                    GetSearchControlByControlName = cmdProg_Spector_CdSearch

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

                Case "cmdOrderSearch" ' "txtOrder"
                    If chkExactRepeat.Checked Then
                        GetSearchElementByControlName = "OEI-HDR-VIEW" ' "OE-HDR-HST-ORD"
                    Else
                        GetSearchElementByControlName = "OEI-ORDER-NO"
                    End If

                Case "txtApplyTo", "cmdApplyToSearch"
                    GetSearchElementByControlName = "OE-HDR-HST-INV"

                Case "txtCus_No", "cmdCustomerSearch"
                    '==============================
                    'Modified May 30, 2017 - T. Louzon - Changing search to add more columns
                    'GetSearchElementByControlName = "AR-CUSTOMER"
                    GetSearchElementByControlName = "OEI-CUSTOMER"
                    '==============================

                Case "txtSlspsn_No", "cmdSalespersonSearch"
                    GetSearchElementByControlName = "AR-SALESPERSON"

                Case "txtJob_No", "cmdProjectSearch"
                    GetSearchElementByControlName = "SY-JOB-CODE"

                Case "txtMfg_Loc", "cmdLocationSearch"
                    GetSearchElementByControlName = "IM-LOCATION"

                Case "txtAr_Terms_Cd", "cmdTermsSearch"
                    GetSearchElementByControlName = "AR-TERMS-CODE"

                Case "txtProfit_Center", "cmdProfitCenterSearch"
                    GetSearchElementByControlName = "SY-ACCOUNT2"

                Case "txtDept", "cmdCostCenterSearch"
                    GetSearchElementByControlName = "SY-ACCOUNT3"

                Case "txtShip_Instruction_1", "cmdInstructionsSearch"
                    GetSearchElementByControlName = "OE-COMMENT-CD-2"

                Case "txtShip_Instruction_2", "cmdCommentsSearch"
                    GetSearchElementByControlName = "OE-COMMENT-CD-2"

                Case "txtBill_To_Country", "cmdBillToCountrySearch"
                    GetSearchElementByControlName = "IM-LANDCODE"

                Case "txtShip_To_Country", "cmdShipToCountrySearch"
                    GetSearchElementByControlName = "IM-LANDCODE"

                Case "txtCus_Alt_Adr_Cd"
                    GetSearchElementByControlName = "AR-ALT-ADDR-COD"

                Case "txtShip_Via_Cd", "cmdShipViaSearch"
                    GetSearchElementByControlName = "AR-SHIP-VIA"

                Case "txtUser_Def_Fld_3", "cmdUser_Def_Fld_3Search"
                    GetSearchElementByControlName = "OEI-CSR-USERS"

                Case "txtUser_Def_Fld_4", "cmdUser_Def_Fld_4Search"
                    GetSearchElementByControlName = "OEI-REPEAT-CNT"

                Case "txtUser_Def_Fld_5", "cmdUser_Def_Fld_5Search"
                    GetSearchElementByControlName = "OEI-REP"
                    'GetSearchElementByControlName = "OEI-SHIP-LINK"

                Case "txtShip_Link", "cmdShip_LinkSearch"
                    GetSearchElementByControlName = "OEI-SHIP-LINK"
                    'GetSearchElementByControlName = "OEI-REP"

                Case "txtShip_To_Name", "cmdShip_To_NameSearch"
                    GetSearchElementByControlName = "OEI-SHIP-ADDR"

                Case "txtProg_Spector_Cd", "cmdProg_Spector_CdSearch"
                    GetSearchElementByControlName = "OEI-PROGRAM"


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

                Case "AR-ALT-ADDR-COD"
                    GetSearchFieldByElement = "ARALTADR_SQL.cus_alt_adr_cd"

                Case "AR-CUSTOMER"
                    GetSearchFieldByElement = "arcusfil_sql.cus_no"

                '==============================
                'Added May 30, 2017 - T. Louzon
                Case "OEI-CUSTOMER"
                    GetSearchFieldByElement = "arcusfil_sql.cus_no"
                '==============================

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

                Case "OEI-ORDER-NO"
                    GetSearchFieldByElement = "OEI_ORDHDR.Cus_No"

                Case "OEI-CSR-USERS"
                    GetSearchFieldByElement = "HUMRES.FullName"

                Case "OEI-CUS-CONTACT"
                    GetSearchFieldByElement = "OEORDHDR_SQL.OE_PO_No"

                Case "OEI-REP"
                    GetSearchFieldByElement = "Exact_Traveler_Rep_Codes.RepCust"

                Case "OEI-SHIP-LINK"
                    GetSearchFieldByElement = "OEORDHDR_SQL.OE_PO_No"

                Case "OEI-PROGRAM"
                    GetSearchFieldByElement = "MDB_CUS_PROG.Spector_Cd"

                Case "SY-ACCOUNT2"
                    GetSearchFieldByElement = "kstdr.kstdrcode"

                Case "SY-ACCOUNT3"
                    GetSearchFieldByElement = "kstpl.kstplcode"

                Case "SY-JOB-CODE"
                    GetSearchFieldByElement = "jobfile_sql.job_no"

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Search buttons click - open search window for element
    Private Sub Search_Element_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _
                cmdBillToCountrySearch.Click, _
                cmdCommentsSearch.Click, cmdDeptSearch.Click, cmdCustomerSearch.Click, _
                cmdInstructionsSearch.Click, cmdLocationSearch.Click, cmdOrderSearch.Click, _
                cmdProfitCenterSearch.Click, cmdProjectSearch.Click, cmdSalespersonSearch.Click, _
                cmdShipToCountrySearch.Click, cmdShipViaSearch.Click, cmdTermsSearch.Click, _
                cmdUser_Def_Fld_5Search.Click, cmdUser_Def_Fld_3Search.Click ' , cmdUser_Def_Fld_4Search.Click

        If blnForceFocus Then Exit Sub

        Dim btnElement As Button

        Try

            If TypeOf eventSender Is Button Then

                btnElement = New Button
                btnElement = DirectCast(eventSender, Button)

                Dim oSearch As New cSearch(GetSearchElementByControlName(btnElement.Name))
                If oSearch.Form.FoundRow <> -1 Then
                    Dim txtElement As Object


                    txtElement = GetElementByControlName(btnElement.Name)

                    '++ID 12.12.2024 concatenate if is the case shipping instruction 1
                    If btnElement.Name = "cmdInstructionsSearch" Then
                        txtElement.Text &= "  " & Trim(oSearch.Form.FoundElementValue)
                    Else
                        txtElement.Text = oSearch.Form.FoundElementValue
                    End If


                    Select Case btnElement.Name
                            Case "cmdUser_Def_Fld_3Search"
                                'ttToolTip.SetToolTip(txtElement, oSearch.Form.TextMatrix(oSearch.Form.FoundRow, 2))
                                ttToolTip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(oSearch.Form.FoundRow).Cells(2).Value)




                            Case Else
                                'ttToolTip.SetToolTip(txtElement, oSearch.Form.TextMatrix(oSearch.Form.FoundRow, 3))
                                ttToolTip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(oSearch.Form.FoundRow).Cells(3).Value)

                        End Select

                        txtElement.focus()

                    End If
                    oSearch.Dispose()
                oSearch = Nothing

                If btnElement.Name = "cmdCustomerSearch" Then txtCus_No_Validated(txtCus_No, New System.EventArgs)

            End If
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Private events - GotFocus ########################################"

    Private Sub txtOrd_Dt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOrd_Dt.GotFocus

        If Trim(txtOrd_Dt.Text.ToString) = "" Then txtOrd_Dt.Text = Date.Now.ToShortDateString

    End Sub



    Private Sub txtInHandsDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInHandsDate.GotFocus

        If Trim(txtInHandsDate.Text.ToString) = "" Then txtInHandsDate.Text = Date.Now.ToShortDateString

    End Sub

    Private Sub txtOrd_Dt_Shipped_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOrd_Dt_Shipped.GotFocus

        If Trim(txtOrd_Dt_Shipped.Text.ToString) = "" Then
            If g_User.Group_Defaults.Contains("ORD_DT_SHIPPED") Then
                If g_User.Group_Defaults.Item("ORD_DT_SHIPPED") = "NOW()" Then
                    txtOrd_Dt_Shipped.Text = Date.Now.ToShortDateString
                Else
                    txtOrd_Dt_Shipped.Text = "A.S.A.P."
                End If
            Else
                txtOrd_Dt_Shipped.Text = "A.S.A.P."
            End If
        End If

        'If g_User.Group_Defaults.Contains("") Then
        '    txtOrd_Dt_Shipped.Text = Trim(g_User.Group_Defaults.Item("").ToString)
        'End If 
    End Sub

    ' Logical sequence get focus to force focus on needed fields
    Private Sub Element_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
                    txtOrd_Dt.GotFocus, cmdDateTP.GotFocus, txtOrd_Dt_Shipped.GotFocus,
                    cmdShipDateTP.GotFocus, txtCus_Alt_Adr_Cd.GotFocus, txtCus_No.GotFocus,
                    cmdCustomerSearch.GotFocus, txtOe_Po_No.GotFocus, txtSlspsn_No.GotFocus,
                    txtJob_No.GotFocus, txtShip_Instruction_1.GotFocus, txtMfg_Loc.GotFocus,
                    txtContactEmail.GotFocus, txtShip_Via_Cd.GotFocus, txtAr_Terms_Cd.GotFocus,
                    txtDiscount_Pct.GotFocus, txtProfit_Center.GotFocus, txtDept.GotFocus,
                    cmdDept_Notes.GotFocus, txtShip_Instruction_2.GotFocus, cmdOrderSearch.GotFocus,
                    cmdCustomerSearch.GotFocus, CmdShipToSearch.GotFocus, cmdSalespersonSearch.GotFocus,
                    cmdProjectSearch.GotFocus, cmdInstructionsSearch.GotFocus,
                    cmdLocationSearch.GotFocus, cmdShipViaSearch.GotFocus, cmdTermsSearch.GotFocus,
                    cmdProfitCenterSearch.GotFocus, cmdDeptSearch.GotFocus, cmdCommentsSearch.GotFocus,
                    txtBill_To_Name.GotFocus, txtBill_To_Addr_1.GotFocus, txtBill_To_Addr_2.GotFocus, txtBill_To_Addr_3.GotFocus, txtBill_To_Addr_4.GotFocus, txtBill_To_City.GotFocus, txtBill_To_State.GotFocus, txtBill_To_Zip.GotFocus,
                    txtShip_To_Name.GotFocus, txtShip_To_Addr_1.GotFocus, txtShip_To_Addr_2.GotFocus, txtShip_To_Addr_3.GotFocus, txtShip_To_Addr_4.GotFocus, txtShip_To_City.GotFocus, txtShip_To_State.GotFocus, txtShip_To_Zip.GotFocus, cmdShip_To_NameSearch.GotFocus,
                    txtBill_To_Country.GotFocus, cmdBillToCountrySearch.GotFocus, txtShip_To_Country.GotFocus, cmdShipToCountrySearch.GotFocus,
                    txtUser_Def_Fld_1.GotFocus, txtEnd_User.GotFocus, cboUser_Def_Fld_2.GotFocus, txtUser_Def_Fld_3.GotFocus, cmdUser_Def_Fld_3Search.GotFocus,
                    txtUser_Def_Fld_4.GotFocus, cmdUser_Def_Fld_4Search.GotFocus, txtUser_Def_Fld_5.GotFocus, cmdUser_Def_Fld_5Search.GotFocus,
                    txtShip_Link.GotFocus, cmdShip_LinkSearch.GotFocus, cmdShip_LinkEdit.GotFocus, cbo_Extra_9.GotFocus,
                    cmdOrderLink.GotFocus, cmdComments.GotFocus, cmdOEFlash.GotFocus, txtProg_Spector_Cd.GotFocus, cmdProg_Spector_CdSearch.GotFocus ' , chkOrderAckSaveOnly.GotFocus
        'ici'
        ' , chkOrderAckSaveOnly.GotFocus

        '++ID 06.17.2020 added new field only for oei_ordheader cbo_Extra_9.GotFocus

        Dim lTabIndex As Integer

        Try
            If TypeOf sender Is xTextBox Then

                Dim txtElement As New xTextBox
                txtElement = DirectCast(sender, xTextBox)
                lTabIndex = txtElement.TabIndex

                Select Case txtElement.Name
                    Case "txtOrd_Dt", "txtOrd_Dt_Shipped", "txtInHandDate"
                        Call oFrmOrder.SetStatusBarPanelText(1, g_Date_Shortcut)
                        'txtElement.BeginInvoke(New myDelegate(AddressOf frmOrder.RefreshStatusBarPanelText))
                        'txtElement.BeginInvoke(New myDelegate(

                    Case "txtOrd_No", "txtApplyTo", "txtCus_No", "txtCus_Alt_Adr_Cd", "txtSlspsn_No",
                         "txtJob_No", "txtMfg_Loc", "txtShip_Via_Cd", "txtAr_Terms_Cd", "txtDept",
                         "txtBill_To_Country", "txtShip_To_Country", "txtProfit_Center"
                        Call oFrmOrder.SetStatusBarPanelText(1, g_Search_Shortcut & g_Notes_Shortcut)
                        'txtElement.BeginInvoke(New myDelegate(AddressOf frmOrder.RefreshStatusBarPanelText))
                        If txtElement.Name = "txtCus_Alt_Adr_Cd" And txtElement.Text.ToString.ToString = "SAME" Then
                            txtElement.Text = ""
                        End If

                    Case "txtShip_Link"
                        Call oFrmOrder.SetStatusBarPanelText(1, g_Search_Shortcut & g_Ship_Links_Shortcut)

                    Case "txtShip_Instruction_1", "txtShip_Instruction_2", "txtUser_Def_Fld_3", "txtUser_Def_Fld_4", "txtUser_Def_Fld_5", "txtShip_To_Name", "txtProg_Spector_Cd"
                        Call oFrmOrder.SetStatusBarPanelText(1, g_Search_Shortcut)
                        'txtElement.BeginInvoke(New myDelegate(AddressOf frmOrder.RefreshStatusBarPanelText))
                         '++ID 06.17.2020 added new field only for oei_ordheader
                    Case "txtProfit_Center", "cboUser_Def_Fld_2", "cbo_Extra_9"
                        Call oFrmOrder.SetStatusBarPanelText(1, " ")

                    Case Else
                        Call oFrmOrder.SetStatusBarPanelText(1, " ")
                        'txtElement.BeginInvoke(New myDelegate(AddressOf frmOrder.RefreshStatusBarPanelText))

                End Select

            ElseIf TypeOf sender Is ComboBox Then

                Dim txtElement As New ComboBox
                txtElement = DirectCast(sender, ComboBox)
                lTabIndex = txtElement.TabIndex

                Call oFrmOrder.SetStatusBarPanelText(1, " ")

            ElseIf TypeOf sender Is CheckBox Then

                Dim txtElement As New CheckBox
                txtElement = DirectCast(sender, CheckBox)
                lTabIndex = txtElement.TabIndex

                Call oFrmOrder.SetStatusBarPanelText(1, " ")

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

    '' Logical sequence get focus to force focus on needed fields
    'Private Sub Element_GotFocus1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
    '    txtProg_Spector_Cd.GotFocus, cmdProg_Spector_CdSearch.GotFocus ' , chkOrderAckSaveOnly.GotFocus

    '    Dim lTabIndex As Integer

    '    Try
    '        If TypeOf sender Is xTextBox Then

    '            Dim txtElement As New xTextBox
    '            txtElement = DirectCast(sender, xTextBox)
    '            lTabIndex = txtElement.TabIndex

    '            Select Case txtElement.Name
    '                Case "txtOrd_Dt", "txtOrd_Dt_Shipped"
    '                    Call oFrmOrder.SetStatusBarPanelText(1, g_Date_Shortcut)
    '                    'txtElement.BeginInvoke(New myDelegate(AddressOf frmOrder.RefreshStatusBarPanelText))
    '                    'txtElement.BeginInvoke(New myDelegate(

    '                Case "txtOrd_No", "txtApplyTo", "txtCus_No", "txtCus_Alt_Adr_Cd", "txtSlspsn_No", _
    '                     "txtJob_No", "txtMfg_Loc", "txtShip_Via_Cd", "txtAr_Terms_Cd", "txtDept", _
    '                     "txtBill_To_Country", "txtShip_To_Country", "txtProfit_Center"
    '                    Call oFrmOrder.SetStatusBarPanelText(1, g_Search_Shortcut & g_Notes_Shortcut)
    '                    'txtElement.BeginInvoke(New myDelegate(AddressOf frmOrder.RefreshStatusBarPanelText))
    '                    If txtElement.Name = "txtCus_Alt_Adr_Cd" And txtElement.Text.ToString.ToString = "SAME" Then
    '                        txtElement.Text = ""
    '                    End If

    '                Case "txtShip_Link"
    '                    Call oFrmOrder.SetStatusBarPanelText(1, g_Search_Shortcut & g_Ship_Links_Shortcut)

    '                Case "txtShip_Instruction_1", "txtShip_Instruction_2", "txtUser_Def_Fld_3", "txtUser_Def_Fld_5", "txtShip_To_Name"
    '                    Call oFrmOrder.SetStatusBarPanelText(1, g_Search_Shortcut)
    '                    'txtElement.BeginInvoke(New myDelegate(AddressOf frmOrder.RefreshStatusBarPanelText))

    '                Case "txtProfit_Center", "cboUser_Def_Fld_2"
    '                    Call oFrmOrder.SetStatusBarPanelText(1, " ")

    '                Case Else
    '                    Call oFrmOrder.SetStatusBarPanelText(1, " ")
    '                    'txtElement.BeginInvoke(New myDelegate(AddressOf frmOrder.RefreshStatusBarPanelText))

    '            End Select

    '        ElseIf TypeOf sender Is ComboBox Then

    '            Dim txtElement As New ComboBox
    '            txtElement = DirectCast(sender, ComboBox)
    '            lTabIndex = txtElement.TabIndex

    '            Call oFrmOrder.SetStatusBarPanelText(1, " ")

    '        ElseIf TypeOf sender Is CheckBox Then

    '            Dim txtElement As New CheckBox
    '            txtElement = DirectCast(sender, CheckBox)
    '            lTabIndex = txtElement.TabIndex

    '            Call oFrmOrder.SetStatusBarPanelText(1, " ")

    '        ElseIf TypeOf sender Is Button Then

    '            Dim txtElement As New Button
    '            txtElement = DirectCast(sender, Button)
    '            lTabIndex = txtElement.TabIndex

    '        End If


    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    '    blnForceFocus = ForceFocus(lTabIndex)

    'End Sub

#End Region


#Region "Private events - KeyDown #########################################"

    Private Sub txtOrd_No_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOrd_No.GotFocus

        txtOrd_No.Enabled = (Trim(txtCus_No.Text) = "")
        oFrmOrder._sbOrder_Panel1.Text = g_Search_Shortcut & g_Notes_Shortcut

    End Sub

    Private Sub cboOrd_Type_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboOrd_Type.GotFocus

        oFrmOrder.SetStatusBarPanelText(1, g_Order_Type_ShortCut)

    End Sub


    ' Gets Key Down, to trigger ESC, F7, F8 and others
    Private Sub Pressed_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
        cboOrd_Type.KeyDown, txtOrd_No.KeyDown, txtCus_No.KeyDown, txtOrd_Dt.KeyDown,
        txtOrd_Dt_Shipped.KeyDown, txtCus_Alt_Adr_Cd.KeyDown, txtOe_Po_No.KeyDown,
        txtSlspsn_No.KeyDown, txtJob_No.KeyDown, txtMfg_Loc.KeyDown,
        txtShip_Via_Cd.KeyDown, txtAr_Terms_Cd.KeyDown, txtDiscount_Pct.KeyDown,
        txtProfit_Center.KeyDown, txtDept.KeyDown, txtShip_Instruction_1.KeyDown,
        txtShip_Instruction_2.KeyDown, txtBill_To_Name.KeyDown, txtBill_To_Addr_1.KeyDown,
        txtBill_To_Addr_2.KeyDown, txtBill_To_Addr_3.KeyDown, txtBill_To_Addr_4.KeyDown, txtBill_To_City.KeyDown, txtBill_To_State.KeyDown, txtBill_To_Zip.KeyDown,
        txtBill_To_Country.KeyDown, txtShip_To_Name.KeyDown, txtShip_To_Addr_1.KeyDown,
        txtShip_To_Addr_2.KeyDown, txtShip_To_Addr_3.KeyDown, txtShip_To_Addr_4.KeyDown, txtShip_To_City.KeyDown, txtShip_To_State.KeyDown, txtShip_To_Zip.KeyDown,
        txtShip_To_Country.KeyDown, txtEnd_User.KeyDown, txtUser_Def_Fld_1.KeyDown, cboUser_Def_Fld_2.KeyDown, cbo_Extra_9.KeyDown,
        txtUser_Def_Fld_3.KeyDown, txtUser_Def_Fld_4.KeyDown, txtUser_Def_Fld_5.KeyDown,
        txtShip_Link.KeyDown, txtProg_Spector_Cd.KeyDown, txtInHandsDate.KeyDown

        '++ID 06.17.2020 added new field only for cbo_extra_9.keydown
        Dim txtElement As New Object

        Try

            '' If Ctl-Number, sets the new tab.
            'If e.Control Then
            '    Select Case e.KeyCode
            '        Case Keys.D1
            '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Header)
            '        Case Keys.D2
            '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CustomerContact)
            '        Case Keys.D3
            '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Lines)
            '        Case Keys.D4
            '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Documents)
            '        Case Keys.D5
            '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.HeaderAll)
            '        Case Keys.D6
            '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Taxes)
            '        Case Keys.D7
            '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Salesperson)
            '        Case Keys.D8
            '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
            '        Case Keys.D9
            '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
            '    End Select

            'End If

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
                    'If txtElement.tabindex < txtCus_No.TabIndex Then
                    '    MsgBox("On supprime tout")
                    'Else
                    '    Dim oReponse As New MsgBoxResult
                    '    oReponse = MsgBox("Save Changes ?", MsgBoxStyle.YesNoCancel, "OE Interface")
                    '    If oReponse = MsgBoxResult.Yes Then
                    '        MsgBox("On save")
                    '    ElseIf oReponse = MsgBoxResult.No Then
                    '        MsgBox("on annule la commande")
                    '    Else
                    '        MsgBox("on fait rien")
                    '    End If
                    'End If

                    ' New order number
                Case Keys.F5
                    If txtElement.name <> "txtOrd_No" Then
                        Exit Sub
                    End If
                    'Call CheckPendingOrder(OrderNoChangeType.NewOrder)

                    'If Trim(txtOrd_No.Text) <> "" Then
                    '    Dim intReponse As MsgBoxResult
                    '    intReponse = MsgBox("Your order will be pending. Continue ?", MsgBoxStyle.YesNoCancel)
                    '    Select Case intReponse
                    '        Case MsgBoxResult.Yes

                    Dim strOrd_No As String = txtOrd_No.Text
                    'If Not (m_oOrder.Ordhead.OEI_Ord_No Is Nothing) Then strOrd_No = m_oOrder.Ordhead.OEI_Ord_No
                    If Trim(strOrd_No) = "" Then
                        m_oOrder = New cOrder
                        txtOrd_No.Text = g_User.NextOrderNumber()
                        If xml_Access = "0" Then
                            cmdXml.Visible = False
                        Else
                            cmdXml.Visible = True
                        End If
                    Else
                        Dim dt As DataTable
                        Dim db As New cDBA
                        dt = db.DataTable("SELECT * FROM OEI_OrdHdr WITH (Nolock) WHERE OEI_Ord_No = '" & m_oOrder.Ordhead.OEI_Ord_No & "' ")
                        If dt.Rows.Count <> 0 Then
                            m_oOrder = New cOrder

                            'Dim txtNextOrderNumber As String = g_User.NextOrderNumber()

                            'If txtNextOrderNumber <> txtOrd_No.Text Then
                            txtOrd_No.Text = g_User.NextOrderNumber()
                            'End If
                            '    Case MsgBoxResult.No
                            '    Case MsgBoxResult.Cancel
                            'End Select
                            '        Else
                            'txtOrd_No.Text = g_User.NextOrderNumber()
                            '        End If
                        End If
                    End If

                Case Keys.F7
                    ' Search
                    Dim oButton As Button
                    oButton = GetSearchControlByControlName(sender.name)
                    If Not (oButton.Name Is Nothing) Then oButton.PerformClick()

                Case Keys.F8
                    ' Notes
                    Dim oButton As Button
                    oButton = GetNotesControlByControlName(sender.name)
                    If Not (oButton.Name Is Nothing) Then oButton.PerformClick()


                Case Keys.F9
                    ' Calendar
                    Dim oButton As Button
                    oButton = GetDateControlByControlName(sender.name)
                    If Not (oButton.Name Is Nothing) Then oButton.PerformClick()

                    'Case Keys.Control
                    '    Select Case e.KeyCode
                    '        Case Keys.D1
                    '            ' oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Header)
                    '        Case Keys.D2
                    '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CustomerContact)
                    '        Case Keys.D3
                    '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Lines)
                    '        Case Keys.D4
                    '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Documents)
                    '        Case Keys.D5
                    '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.HeaderAll)
                    '        Case Keys.D6
                    '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Taxes)
                    '        Case Keys.D7
                    '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Salesperson)
                    '        Case Keys.D8
                    '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
                    '        Case Keys.D9
                    '            oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
                    '    End Select

            End Select

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    '' Gets Key Down, to trigger ESC, F7, F8 and others
    'Private Sub Pressed_KeyDown1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
    '    txtProg_Spector_Cd.KeyDown

    '    Dim txtElement As New Object

    '    Try

    '        If TypeOf sender Is xTextBox Then

    '            txtElement = New xTextBox
    '            txtElement = DirectCast(sender, xTextBox)

    '        ElseIf TypeOf sender Is ComboBox Then

    '            txtElement = New ComboBox
    '            txtElement = DirectCast(sender, ComboBox)

    '        End If

    '        Select Case e.KeyValue

    '            ' New order number
    '            Case Keys.F5
    '                If txtElement.name <> "txtOrd_No" Then
    '                    Exit Sub
    '                End If

    '                Dim strOrd_No As String = txtOrd_No.Text
    '                'If Not (m_oOrder.Ordhead.OEI_Ord_No Is Nothing) Then strOrd_No = m_oOrder.Ordhead.OEI_Ord_No
    '                If Trim(strOrd_No) = "" Then
    '                    m_oOrder = New cOrder
    '                    txtOrd_No.Text = g_User.NextOrderNumber()
    '                Else
    '                    Dim dt As DataTable
    '                    Dim db As New cDBA
    '                    dt = db.DataTable("SELECT * FROM OEI_OrdHdr WITH (Nolock) WHERE OEI_Ord_No = '" & m_oOrder.Ordhead.OEI_Ord_No & "' ")
    '                    If dt.Rows.Count <> 0 Then
    '                        m_oOrder = New cOrder

    '                        txtOrd_No.Text = g_User.NextOrderNumber()
    '                    End If
    '                End If

    '            Case Keys.F7
    '                ' Search
    '                Dim oButton As Button
    '                oButton = GetSearchControlByControlName(sender.name)
    '                If Not (oButton.Name Is Nothing) Then oButton.PerformClick()

    '            Case Keys.F8
    '                ' Notes
    '                Dim oButton As Button
    '                oButton = GetNotesControlByControlName(sender.name)
    '                If Not (oButton.Name Is Nothing) Then oButton.PerformClick()


    '            Case Keys.F9
    '                ' Calendar
    '                Dim oButton As Button
    '                oButton = GetDateControlByControlName(sender.name)
    '                If Not (oButton.Name Is Nothing) Then oButton.PerformClick()

    '        End Select

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

#End Region


#Region "Private events - LostFocus #######################################"

    ' Sets field color to white or red when data not found in DB
    Private Sub Element_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
                    txtCus_No.LostFocus, txtCus_Alt_Adr_Cd.LostFocus,
                    txtSlspsn_No.LostFocus, txtJob_No.LostFocus, txtShip_Via_Cd.LostFocus,
                    txtAr_Terms_Cd.LostFocus, txtProfit_Center.LostFocus, txtDept.LostFocus,
                    txtBill_To_Country.LostFocus, txtShip_To_Country.LostFocus, txtMfg_Loc.LostFocus,
                    txtUser_Def_Fld_5.LostFocus, txtShip_Link.LostFocus, txtProg_Spector_Cd.LostFocus,
                    txtShip_To_State.LostFocus

        'txtShip_Instruction_1.LostFocus, txtShip_Instruction_2.LostFocus, txtUser_Def_Fld_4.LostFocus, 

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

            ' If record not on file, then set field red and set back focus to it.
            If ttToolTip.GetToolTip(txtElement) = "Record not on file." Then
                txtElement.BackColor = Color.FromArgb(255, 192, 192)
                ForceFocus(txtElement)
            Else
                txtElement.BackColor = Color.White
            End If

            If txtElement.name = "txtCus_Alt_Adr_Cd" Then
                m_oOrder.Ordhead.GetAlternateAddress(txtCus_Alt_Adr_Cd.Text.ToString)

                If m_oOrder.Ordhead.Cus_Alt_Adr_Cd = "SAME" Then
                    txtCus_Alt_Adr_Cd.Text = "SAME"
                End If

                If txtShip_To_Name.ForeColor = Color.Black Then
                    txtShip_To_Name.Text = m_oOrder.Ordhead.Ship_To_Name
                    txtShip_To_Addr_1.Text = m_oOrder.Ordhead.Ship_To_Addr_1
                    txtShip_To_Addr_2.Text = m_oOrder.Ordhead.Ship_To_Addr_2
                    txtShip_To_Addr_3.Text = m_oOrder.Ordhead.Ship_To_Addr_3
                    txtShip_To_Addr_4.Text = m_oOrder.Ordhead.Ship_To_Addr_4
                    txtShip_To_City.Text = m_oOrder.Ordhead.Ship_To_City
                    txtShip_To_State.Text = m_oOrder.Ordhead.Ship_To_State
                    txtShip_To_Zip.Text = m_oOrder.Ordhead.Ship_To_Zip
                    txtShip_To_Country.Text = m_oOrder.Ordhead.Ship_To_Country


                    If m_oOrder.Ordhead.IsARepOrder() Then
                        If txtShip_To_Name.Text = txtBill_To_Name.Text And txtShip_To_Addr_1.Text = txtBill_To_Addr_1.Text Then
                            If txtShip_Instruction_1.Text = String.Empty Then
                                txtShip_Instruction_1.Text = "NO CHARGE FOR FREIGHT"
                            ElseIf txtShip_Instruction_2.Text = String.Empty And Trim(txtShip_Instruction_1.Text) <> "NO CHARGE FOR FREIGHT" Then
                                txtShip_Instruction_2.Text = "NO CHARGE FOR FREIGHT"
                                'Else
                                '    MsgBox("NO CHARGE FOR FREIGHT")
                            End If

                        End If

                    End If

                End If

                'ElseIf txtElement.name = "txtShip_Via_Cd" And txtElement.text <> String.Empty And m_oOrder.Ordhead.Customer.ClassificationId <> String.Empty Then

                'moved to frmOrder.sstOrder.Click 2960
                '   Dim _mestext As String = ""
                '_mestext = spec_ship_charge(txtElement.text, m_oOrder.Ordhead.Customer.ClassificationId)

                ''  If txtElement.Text = "72" And m_oOrder.Ordhead.Customer.ClassificationId = "6ST" And mglobal_changed_shipvia <> True Then
                'If mglobal_changed_shipvia <> True And _mestext <> String.Empty Then
                '    mglobal_changed_shipvia = True
                '    MsgBox(_mestext)
                'Else
                '    mglobal_changed_shipvia = False
                'End If

            ElseIf (txtElement.name = "txtShip_To_Country" Or txtElement.name = "txtShip_To_State") Then

                '++ID 02.20.2023
                'If Trim(txtShip_To_Country.Text) <> String.Empty And Trim(txtShip_To_State.Text) <> String.Empty Then

                '    m_oOrder.Ordhead.Mfg_Loc = LocationBy_CountryState(txtShip_To_Country.Text, txtShip_To_State.Text)

                '    Dim _changLoc As String = ""
                '    _changLoc = Trim(txtMfg_Loc.Text)

                '    If Trim(txtMfg_Loc.Text) <> Trim(m_oOrder.Ordhead.Mfg_Loc) Then
                '        txtMfg_Loc.Text = m_oOrder.Ordhead.Mfg_Loc
                '        MsgBox("Take care , Order Location Changed from : Loc : " & _changLoc & " to Loc : " & Trim(m_oOrder.Ordhead.Mfg_Loc) _
                '               & vbCrLf & "Order lines will be changed to Loc : " & Trim(m_oOrder.Ordhead.Mfg_Loc))
                '    End If


                'End If

            ElseIf txtElement.name = "txtCus_No" Then

                '    Call RPSSP(Trim(txtCus_No.Text))

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public Function spec_ship_charge(ByVal m_ShipViaCd As String, ByVal m_classification_id As String) As String

    '    spec_ship_charge = ""

    '    Try

    '        Dim db As New cODBC
    '        Dim dt As DataTable
    '        Dim sql As String

    '        sql = "Select SPEC_INSTRUCTION,shipviacd,levelstar from OEI_ITEM_SPEC_INSTRUCTION " _
    '            & " where shipviacd =  '" & Trim(m_ShipViaCd) & "' and levelstar = '" & Trim(m_classification_id) & "'"

    '        dt = db.DataTable(sql)

    '        If dt.Rows.Count <> 0 Then

    '            Return dt.Rows(0).Item("SPEC_INSTRUCTION").ToString

    '        End If

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Function

    Public Function LocationBy_CountryState(ByVal m_ShipCountry As String, ByVal m_ShipState As String) As String

        LocationBy_CountryState = "1"

        Try

            Dim db As New cODBC
            Dim dt As DataTable
            Dim sql As String

            sql = "Select * from MDB_CFG_ENUM   where ENUM_CAT =  'COUNTRY_STATE' " _
                & " AND ENUM_SUB_CAT = '" & Trim(m_ShipCountry) & "' and ENUM_VALUE = '" & Trim(m_ShipState) & "'"

            dt = db.DataTable(sql)

            If dt.Rows.Count <> 0 Then

                Return "US1"

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function


#End Region


#Region "Private events - TextChanged #####################################"

    ' Autocomplete fields when needed, change fore colors and fill complementary labels
    Private Sub Element_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles txtCus_No.TextChanged, txtSlspsn_No.TextChanged, txtJob_No.TextChanged,
            txtShip_Instruction_1.TextChanged, txtMfg_Loc.TextChanged, txtShip_Instruction_2.TextChanged,
            txtShip_Via_Cd.TextChanged, txtAr_Terms_Cd.TextChanged, txtProfit_Center.TextChanged,
            txtDept.TextChanged, txtBill_To_Country.TextChanged, txtShip_To_Country.TextChanged,
            txtUser_Def_Fld_3.TextChanged, txtUser_Def_Fld_5.TextChanged,
            txtShip_Link.TextChanged, txtProg_Spector_Cd.TextChanged ' , txtUser_Def_Fld_4.TextChanged
        ' Do not put txtShip_To_Name here as it would cause problems loading...
        Try
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

            ElseIf TypeOf sender Is xTextBox Then
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
                'If Not (lblDescriptionLabel Is Nothing) Then
                '    lblDescriptionLabel.Text = ""
                'End If
                ttToolTip.SetToolTip(txtElement, "")
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

            If oSearch.Form.dgvSearch.Rows.Count > 2 And Trim(RTrim(sender.text.ToString.ToUpper)) <> Trim(RTrim(oSearch.Form.FoundElementIndex.ToString)) Then

                If Trim(txtElement.text).Length >= 1 Then
                    If TypeOf sender Is xTextBox Then
                        If txtElement.AutoCompleteMode <> AutoCompleteMode.None Then
                            acscAutoComplete = New AutoCompleteStringCollection()
                            For lPos As Integer = 1 To oSearch.Form.dgvSearch.Rows.Count - 1
                                acscAutoComplete.Add(oSearch.Form.DataGrid.Rows(lPos - 1).Cells(2).Value) ' oSearch.Form.TextMatrix(lPos, 2))
                            Next lPos

                            SyncLock txtElement.AutoCompleteCustomSource.SyncRoot
                                txtElement.AutoCompleteCustomSource = acscAutoComplete
                            End SyncLock
                            acscAutoComplete = Nothing
                        End If
                    Else
                        ttToolTip.SetToolTip(txtElement, "Record not on file.")
                    End If
                    'ttToolTip.SetToolTip(txtElement, "Record not on file.")
                End If
                '            ElseIf oSearch.Form.FoundRow = 1 And Trim(RTrim(sender.text.ToString.ToUpper)) = Trim(RTrim(oSearch.Form.FoundElementValue)).ToUpper Then ' oSearch.Form.fgSearch.Rows = 2 And oSearch.Form.FoundRow = 1 Then
            ElseIf oSearch.Form.FoundRow = 1 And Trim(RTrim(sender.text.ToString.ToUpper)) = Trim(RTrim(oSearch.Form.FoundElementIndex)).ToUpper Then ' oSearch.Form.fgSearch.Rows = 2 And oSearch.Form.FoundRow = 1 Then

                ' Be careful to always check if the value is the same, else we will enter in a deadlock
                Select Case txtElement.name
                    Case "txtUser_Def_Fld_3"
                        ttToolTip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(0).Cells(2).Value) 'TextMatrix(oSearch.Form.FoundRow, 2))
                    Case "txtProg_Spector_Cd"
                        ttToolTip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(0).Cells(2).Value) 'TextMatrix(oSearch.Form.FoundRow, 2))
                    Case Else
                        ttToolTip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(0).Cells(3).Value) ' oSearch.Form.TextMatrix(oSearch.Form.FoundRow, 3))
                End Select

                If Not (lblDescriptionLabel Is Nothing) Then
                    lblDescriptionLabel.Text = oSearch.Form.DataGrid.Rows(0).Cells(3).Value ' oSearch.Form.TextMatrix(oSearch.Form.FoundRow, 3)
                End If

                txtElement.backcolor = Color.White

            ElseIf oSearch.Form.FoundRow = 1 And Trim(sender.text.ToString.ToUpper) = Mid(Trim(oSearch.Form.FoundElementIndex).ToUpper, 1, Len(Trim(sender.text.ToString.ToUpper))) Then ' oSearch.Form.fgSearch.Rows = 2 And oSearch.Form.FoundRow = 1 Then

                ' Be careful to always check if the value is the same, else we will enter in a deadlock
                If txtElement.AutoCompleteMode <> AutoCompleteMode.None Then
                    acscAutoComplete = New AutoCompleteStringCollection()
                    acscAutoComplete.Add(Trim(oSearch.Form.FoundElementValue))
                    SyncLock txtElement.AutoCompleteCustomSource.SyncRoot
                        txtElement.AutoCompleteCustomSource = acscAutoComplete
                    End SyncLock
                    acscAutoComplete = Nothing
                End If
                'txtDept.AutoCompleteCustomSource.
                'acscAutoComplete = Nothing
            Else

                ttToolTip.SetToolTip(txtElement, "Record not on file.")
                'txtElement.backcolor = Color.FromArgb(255, 192, 192)
                If Not (lblDescriptionLabel Is Nothing) Then
                    lblDescriptionLabel.Text = ""
                End If

            End If

            oSearch.Dispose()
            oSearch = Nothing

            'If blnSetTags Then txtElement.Tag = txtElement.Text.ToString.ToUpper
            If blnSetTags Then txtElement.OldValue = txtElement.Text.ToString.ToUpper

            'txtElement.forecolor = IIf(txtElement.tag <> txtElement.text.ToString.ToUpper And Not blnLoading, Color.Blue, Color.Black)
            txtElement.forecolor = IIf(txtElement.OldValue <> txtElement.text.ToString.ToUpper And Not blnLoading, Color.Blue, Color.Black)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub


    ' Autocomplete fields when needed, change fore colors and fill complementary labels
    Private Sub txtCus_Alt_Adr_Cd_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles txtCus_Alt_Adr_Cd.TextChanged


        Try
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

            ElseIf TypeOf sender Is xTextBox Then
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

            'If Not (txtElement.name = "txtCus_Alt_Adr_Cd" And (txtElement.text = "SAME" Or txtElement.text = "")) Then
            'If Not (txtElement.name = "txtCus_Alt_Adr_Cd" And (txtElement.text = "SAME")) Then
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

                If Trim(txtElement.text).Length >= 1 Then
                    If TypeOf sender Is xTextBox Then
                        If txtElement.AutoCompleteMode <> AutoCompleteMode.None Then
                            acscAutoComplete = New AutoCompleteStringCollection()
                            For lPos As Integer = 1 To oSearch.Form.dgvSearch.Rows.Count - 1
                                acscAutoComplete.Add(oSearch.Form.DataGrid.Rows(lPos - 1).Cells(2).Value) ' oSearch.Form.TextMatrix(lPos, 2))
                            Next lPos

                            SyncLock txtElement.AutoCompleteCustomSource.SyncRoot
                                txtElement.AutoCompleteCustomSource = acscAutoComplete
                            End SyncLock
                            acscAutoComplete = Nothing
                        End If
                        'acscAutoComplete = Nothing
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
                        ttToolTip.SetToolTip(txtElement, "Record not on file.")
                    End If
                    ttToolTip.SetToolTip(txtElement, "Record not on file.")
                End If
            ElseIf oSearch.Form.FoundRow = 1 And Trim(RTrim(sender.text.ToString.ToUpper)) = Trim(RTrim(oSearch.Form.FoundElementValue)).ToUpper Then ' oSearch.Form.fgSearch.Rows = 2 And oSearch.Form.FoundRow = 1 Then

                ' Be careful to always check if the value is the same, else we will enter in a deadlock
                'If oSearch.Form.FoundElementValue <> txtElement.text Then
                '    txtElement.text = oSearch.Form.FoundElementValue
                'End If
                'txtElement.AutoCompleteCustomSource = Nothing
                Select Case txtElement.name
                    Case "txtUser_Def_Fld_3"
                        ttToolTip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(0).Cells(2).Value) ' TextMatrix(oSearch.Form.FoundRow, 2))
                    Case Else
                        ttToolTip.SetToolTip(txtElement, oSearch.Form.DataGrid.Rows(0).Cells(3).Value) ' oSearch.Form.TextMatrix(oSearch.Form.FoundRow, 3))
                End Select
                If Not (lblDescriptionLabel Is Nothing) Then
                    lblDescriptionLabel.Text = oSearch.Form.DataGrid.Rows(0).Cells(3).Value ' oSearch.Form.TextMatrix(oSearch.Form.FoundRow, 3)
                End If
                txtElement.backcolor = Color.White
            ElseIf oSearch.Form.FoundRow = 1 And Trim(sender.text.ToString.ToUpper) = Mid(Trim(oSearch.Form.FoundElementValue).ToUpper, 1, Len(Trim(sender.text.ToString.ToUpper))) Then ' oSearch.Form.fgSearch.Rows = 2 And oSearch.Form.FoundRow = 1 Then

                ' Be careful to always check if the value is the same, else we will enter in a deadlock
                'If oSearch.Form.FoundElementValue <> txtElement.text Then
                '    txtElement.text = oSearch.Form.FoundElementValue
                'End If
                'txtElement.AutoCompleteCustomSource = Nothing
                If txtElement.AutoCompleteMode <> AutoCompleteMode.None Then
                    acscAutoComplete = New AutoCompleteStringCollection()
                    acscAutoComplete.Add(Trim(oSearch.Form.FoundElementValue))
                    SyncLock txtElement.AutoCompleteCustomSource.SyncRoot
                        txtElement.AutoCompleteCustomSource = acscAutoComplete
                    End SyncLock
                    acscAutoComplete = Nothing
                End If
                'txtDept.AutoCompleteCustomSource.
                'acscAutoComplete = Nothing
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

            'If txtElement.name = "txtCus_Alt_Adr_Cd" Then

            'm_oOrder.Ordhead.GetAlternateAddress()
            'End If

            oSearch.Dispose()
            oSearch = Nothing

            'End If

            'If blnSetTags Then txtElement.Tag = txtElement.Text.ToString.ToUpper
            If blnSetTags Then txtElement.OldValue = txtElement.Text.ToString.ToUpper

            'txtElement.forecolor = IIf(txtElement.tag <> txtElement.text.ToString.ToUpper And Not blnLoading, Color.Blue, Color.Black)
            txtElement.forecolor = IIf(txtElement.OldValue <> txtElement.text.ToString.ToUpper And Not blnLoading, Color.Blue, Color.Black)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Change fore colors on fields with no search
    Private Sub Element_TextChangedNoSearch(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        txtOrd_No.TextChanged, txtOrd_Dt.TextChanged, txtOrd_Dt_Shipped.TextChanged, _
        txtOe_Po_No.TextChanged, txtDiscount_Pct.TextChanged, txtBill_To_Name.TextChanged, _
        txtBill_To_Addr_1.TextChanged, txtBill_To_Addr_2.TextChanged, txtBill_To_Addr_3.TextChanged, _
        txtBill_To_Addr_4.TextChanged, txtBill_To_City.KeyDown, txtBill_To_State.KeyDown, txtBill_To_Zip.KeyDown, txtShip_To_Name.TextChanged, txtShip_To_Addr_1.TextChanged, _
        txtShip_To_Addr_2.TextChanged, txtShip_To_Addr_3.TextChanged, txtShip_To_Addr_4.TextChanged, txtShip_To_City.KeyDown, txtShip_To_State.KeyDown, txtShip_To_Zip.KeyDown, _
        txtEnd_User.TextChanged, txtUser_Def_Fld_1.TextChanged, txtUser_Def_Fld_3.TextChanged, txtInHandsDate.TextChanged, txtUser_Def_Fld_4.TextChanged, txtShip_To_Country.TextChanged,
        txtShip_To_City.TextChanged, txtShip_To_State.TextChanged, txtShip_To_Zip.TextChanged

        'uppercase specific fields
        Dim elem As xTextBox = CType(sender, xTextBox)
        Dim selStart As Integer = elem.SelectionStart

        If elem.Name = "txtUser_Def_Fld_4" OrElse elem.Name = "txtOe_Po_No" OrElse elem.Name = "txtShip_To_Name" OrElse elem.Name = "txtShip_To_Country" _
            OrElse elem.Name = "txtShip_To_Addr_1" OrElse elem.Name = "txtShip_To_Addr_2" OrElse elem.Name = "txtShip_To_Addr_3" _
            OrElse elem.Name = "txtShip_To_Addr_4" OrElse elem.Name = "txtShip_To_City" OrElse elem.Name = "txtShip_To_State" OrElse elem.Name = "txtShip_To_Zip" Then
            elem.Text = elem.Text.ToUpper()

            If elem.SelectionLength = 0 Then
                elem.SelectionStart = selStart 'elem.Text.Length + 1
            End If
        End If



        If elem.Name = "txtShip_To_Country" Or elem.Name = "txtShip_To_State" Then



            '++ID 02.20.2023
            'closing US1 Location, no more need to assigne tpo US11 vegas production ticket #40743
            'If Trim(txtShip_To_Country.Text) <> String.Empty And Trim(txtShip_To_State.Text) <> String.Empty Then
            '    If Trim(txtShip_To_Country.Text).Length >= 2 And Trim(txtShip_To_State.Text).Length >= 2 Then

            '        m_oOrder.Ordhead.Mfg_Loc = LocationBy_CountryState(txtShip_To_Country.Text, txtShip_To_State.Text)

            '        '     If txtMfg_Loc.Text <> m_oOrder.Ordhead.Mfg_Loc Then

            '        Dim _changLoc As String = ""
            '        _changLoc = Trim(txtMfg_Loc.Text)


            '        If Trim(txtMfg_Loc.Text) <> Trim(m_oOrder.Ordhead.Mfg_Loc) Then

            '            txtMfg_Loc.Text = m_oOrder.Ordhead.Mfg_Loc

            '            If Trim(m_oOrder.Ordhead.Mfg_Loc) = "US1" Then

            '                m_oOrder.Ordhead.User_Def_Fld_2 = "US1"
            '                cboUser_Def_Fld_2.Text = "US1"

            '            Else
            '                m_oOrder.Ordhead.User_Def_Fld_2 = "CAD"
            '                cboUser_Def_Fld_2.Text = "CAD"

            '            End If

            '            MsgBox("Take care , Order Location Changed from : Loc : " & _changLoc & " to Loc : " & Trim(m_oOrder.Ordhead.Mfg_Loc) _
            '               & vbCrLf & "Order lines will be changed to Loc : " & Trim(m_oOrder.Ordhead.Mfg_Loc))
            '        End If

            '        '  End If
            '    End If




            'End If


        End If



        Dim txtElement As New Object

        Try
            If TypeOf sender Is xTextBox Then
                txtElement = New xTextBox
                txtElement = DirectCast(sender, xTextBox)

                Dim txtAutoComplete As New xTextBox
                txtAutoComplete = DirectCast(sender, xTextBox)

            ElseIf TypeOf sender Is ComboBox Then

                txtElement = New ComboBox
                txtElement = DirectCast(sender, ComboBox)

            End If

            'If blnSetTags Then txtElement.Tag = txtElement.Text
            If blnSetTags Then txtElement.OldValue = txtElement.Text

            'txtElement.forecolor = IIf(txtElement.tag <> txtElement.text And Not blnLoading, Color.Blue, Color.Black)
            txtElement.forecolor = IIf(txtElement.OldValue <> txtElement.text And Not blnLoading, Color.Blue, Color.Black)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Private events - Validating ######################################"

    'Checks if Ord_dt is a valid date, in an existing period, and period not closed.
    Private Sub txtOrd_Dt_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtOrd_Dt.Validating

        If txtOrd_Dt.Text = "" Then Exit Sub
        'If txtInHandsDate.Text = "" Then Exit Sub ' Don't need that, only validates Ord_Dt is valid.

        Try
            ' Checks for valid date format
            If Not IsDate(txtOrd_Dt.Text) Then
                e.Cancel = True
                Throw New OEException(OEError.Invalid_Date_Format)
            End If

            ' Checks if date entered is in a valid period.
            If Not m_oOrder.Validation.AccountingPeriodExist(CDate(txtOrd_Dt.Text).Year, CDate(txtOrd_Dt.Text).Month) Then
                e.Cancel = True
                Throw New OEException(OEError.No_Accounting_Period_Record_Found)
            End If

            ' Checks if period is a closed period.
            'If m_oOrder.Validation.DateIsOutsideOfValidTrxDates(CDate(txtOrd_Dt.Text).Year, CDate(txtOrd_Dt.Text).Month, "400") Then
            '    e.Cancel = True
            '    Dim strMessage As String = "Entry not allowed " & m_oOrder.Validation.GetPeriodDates(CDate(txtOrd_Dt.Text).Year, CDate(txtOrd_Dt.Text).Month)
            '    Throw New OEException(OEError.Date_Outside_Valid_Trx_Dates, strMessage)
            'End If

        Catch er As OEException
            MsgBox(er.Message)
        End Try

    End Sub

    'Checks if Ord_dt is a valid date, in an existing period, and period not closed.
    Private Sub txtInHandsDt_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtInHandsDate.Validating

        If txtInHandsDate.Text = "" Then Exit Sub

        Try
            ' Checks for valid date format
            If Not IsDate(txtInHandsDate.Text) Then
                e.Cancel = True
                Throw New OEException(OEError.Invalid_Date_Format)
            End If

        Catch er As OEException
            MsgBox(er.Message)
        End Try

    End Sub

    ' Checks if Ord_Dt_Shipped is a valid date.
    Private Sub txtOrd_Dt_Shipped_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtOrd_Dt_Shipped.Validating

        If txtOrd_Dt_Shipped.Text = "" Or txtOrd_Dt_Shipped.Text = "A.S.A.P." Then Exit Sub

        Try
            If Not IsDate(txtOrd_Dt_Shipped.Text) Then
                e.Cancel = True
                Throw New OEException(OEError.Invalid_Date_Format)
            End If

        Catch er As OEException
            MsgBox(er.Message)
        End Try

    End Sub


    ' Checks if PO entered already exists for same customer.
    Private Sub txtOe_Po_No_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtOe_Po_No.Validating

        If Trim(txtOe_Po_No.Text).Length = 0 Then Exit Sub

        Try
            ' Checks if PO entered already exists for same customer.
            Call m_oOrder.Validation.POExists(txtCus_No.Text, txtOe_Po_No.Text)

        Catch oe_er As OEException
            Dim mbrReponse As New MsgBoxResult
            mbrReponse = MsgBox(oe_er.Message, MsgBoxStyle.YesNo)
            If mbrReponse = MsgBoxResult.No Then
                e.Cancel = True
                txtOe_Po_No.Text = txtOe_Po_No.OldValue
            End If
        End Try

    End Sub

    ' Checks if Discount Pct entered between 0 and 100
    Private Sub txtDiscount_Pct_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDiscount_Pct.Validating

        If Trim(txtDiscount_Pct.Text).Length = 0 Then Exit Sub

        Try
            Dim dblValeur As Double
            dblValeur = CDbl(txtDiscount_Pct.Text)

            ' Checks if Discount Pct entered between 0 and 100
            If dblValeur > 100 Then
                Throw New OEException(OEError.Disc_Pct_GT_100_Pct)
            ElseIf dblValeur < 0 Then
                Throw New OEException(OEError.Disc_Pct_LT_0_Pct)
            End If

        Catch oe_er As OEException
            e.Cancel = True
            MsgBox(oe_er.Message)
            txtDiscount_Pct.Text = txtDiscount_Pct.OldValue
        End Try

    End Sub

#End Region


#Region "Control private events ###########################################"

    Private Sub cboCustomerNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCustomerNo.GotFocus

        If blnForceFocus Then Exit Sub

        Try

            If cboCustomerNo.Items.Count = 0 Then
                Dim oSearch As New cSearch("AR-CUSTOMER", False, True)
                For lPos As Integer = 0 To oSearch.Form.DataGrid.Rows.Count - 1
                    cboCustomerNo.Items.Add(oSearch.Form.DataGrid.Rows(lPos - 1).Cells(2).Value) ' oSearch.Form.TextMatrix(lPos, 2))
                Next
                oSearch.Dispose()
                oSearch = Nothing
            End If

            'frmOrder._sbOrder_Panel1.Text = g_Search_Shortcut & g_Notes_Shortcut

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Checks if a customer is valid, if so then sets default values for it.
    Private Sub txtCus_No_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCus_No.Validated

        Try

            'If txtCus_No.Text <> txtCus_No.Tag Then
            'If Trim(txtCus_No.Text) <> Trim(txtCus_No.OldValue) Then
            If Trim(txtCus_No.Text) <> Trim(m_oOrder.Ordhead.Cus_No) Then
                If sender.backcolor = Color.White Then
                    Call SetDefaultCustomerValues()

                End If

                'Call Function to load REP SSP 
                '   If Trim(txtCus_No.Text) <> "" Then
                Call RPSSP()
                '   End If

            End If

            'Select Case m_oOrder.Ordhead.Customer.Status
            '    Case "A", "S" ' Passive and active, ok to proceed
            '    Case "B" ' Blocked, show message. On message if not OK, restart order.
            '        Dim mbResult As New MsgBoxResult
            '        mbResult = MsgBox(New OEExceptionMessage(OEError.Cust_On_Credit_Hold).Message, MsgBoxStyle.YesNo, "Error.")
            '        If mbResult = MsgBoxResult.No Then
            '            ' Cancel the order
            '            MsgBox("Cancel the order")
            '        End If
            '    Case Else '"E" - treat as inactive
            '        MsgBox("Customer is inactive.", MsgBoxStyle.OkOnly, "Error.")
            '        txtCus_No.Text = String.Empty
            'End Select

            'txtCus_No.Tag = txtCus_No.Text
            txtCus_No.OldValue = txtCus_No.Text

            If m_oOrder.Is_White_Flag() Then
                txtUser_Def_Fld_1.Text = "WHITE GLOVE"
                RaiseEvent White_Flag_Customer(True, New System.EventArgs())
            Else
                '++ID 7.10.2018 added elese because if client is white glove and after what change another client without white g. value rest white 
                'else for bring false
                txtUser_Def_Fld_1.Text = ""
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    ' Sets the order number white or pink, is a required field.
    Private Sub txtOrd_No_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOrd_No.TextChanged

        Try
            If String.IsNullOrEmpty(e.ToString) Then
                txtOrd_No.BackColor = Color.FromArgb(255, 192, 192)
            Else
                txtOrd_No.BackColor = Color.White
                'Call CheckPendingOrder()
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Private Z other functions ########################################"

    Private Sub CheckPendingOrder(ByVal pOrderNoChange As OrderNoChangeType)

        Try
            If Trim(txtCus_No.Text) <> "" Then
                Dim intReponse As MsgBoxResult
                intReponse = MsgBox("Your order will be pending. Continue ?", MsgBoxStyle.YesNo)
                Select Case intReponse
                    Case MsgBoxResult.Yes
                        Call Save()

                    Case MsgBoxResult.No

                End Select
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

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
            ' Set focus to Order number if not entered
            If piTabIndex >= txtOrd_Dt.TabIndex And (Trim(txtOrd_No.Text) = "" Or txtOrd_No.BackColor = Color.FromArgb(255, 192, 192)) Then
                txtOrd_No.Focus()
                txtOrd_No.BackColor = Color.FromArgb(255, 192, 192)
                ttToolTip.SetToolTip(txtOrd_No, "Data must be entered here.")
                ForceFocus = True

                Exit Function

            End If

            ' if no date has been entered and click after, then enter today as date
            If piTabIndex >= txtOrd_Dt_Shipped.TabIndex And (txtOrd_Dt.Text = "") Then
                txtOrd_Dt.Text = Date.Now.ToShortDateString
            End If

            '' if no ship date has been entered and click after, then enter today as date
            'If piTabIndex >= txtCus_No.TabIndex And (txtOrd_Dt_Shipped.Text = "") Then
            '    If Trim(txtOrd_Dt_Shipped.Text.ToString) = "" Then
            '        If g_User.Group_Defaults.Contains("ORD_DT_SHIPPED") Then
            '            If g_User.Group_Defaults.Item("ORD_DT_SHIPPED") = "NOW()" Then
            '                txtOrd_Dt_Shipped.Text = Date.Now.ToShortDateString
            '            Else
            '                txtOrd_Dt_Shipped.Text = "A.S.A.P."
            '            End If
            '        Else
            '            txtOrd_Dt_Shipped.Text = "A.S.A.P."
            '        End If
            '    End If
            '    'txtOrd_Dt_Shipped.Text = "A.S.A.P." ' txtOrd_Dt.Text ' Date.Now.ToShortDateString
            'End If

            If piTabIndex >= txtCus_No.TabIndex And (txtOrd_Dt_Shipped.Text = "") Then
                txtOrd_Dt_Shipped.Text = "A.S.A.P." ' txtOrd_Dt.Text ' Date.Now.ToShortDateString
            End If

            ' Set focus to Order number if not entered
            If piTabIndex >= cmdCus_No_Notes.TabIndex And (Trim(txtCus_No.Text) = "" Or txtCus_No.BackColor = Color.FromArgb(255, 192, 192)) Then
                txtCus_No.Focus()
                txtCus_No.BackColor = Color.FromArgb(255, 192, 192)
                ttToolTip.SetToolTip(txtCus_No, "Data must be entered here.")
                ForceFocus = True

                Exit Function

            End If

            ' If no customer has been selected, then you can't select a ship to address.
            If piTabIndex >= txtCus_Alt_Adr_Cd.TabIndex And piTabIndex <= CmdShipToSearch.TabIndex And (txtCus_No.Text = "" Or txtCus_No.BackColor = Color.FromArgb(255, 192, 192)) Then

                txtCus_No.Focus()
                txtCus_No.BackColor = Color.FromArgb(255, 192, 192)
                ttToolTip.SetToolTip(txtCus_No, "Data must be entered here.")
                ForceFocus = True

                Exit Function

            End If

            If piTabIndex >= txtOe_Po_No.TabIndex Then

                ' Si le ship to est vide, on entre celui par defaut.
                If Not (txtCus_No.Text = String.Empty Or txtCus_No.BackColor <> Color.White) Then

                    If Trim(txtShip_To_Name.Text) = String.Empty Then

                        blnSetTags = True

                        txtShip_To_Name.Text = txtBill_To_Name.Text
                        txtShip_To_Addr_1.Text = txtBill_To_Addr_1.Text
                        txtShip_To_Addr_2.Text = txtBill_To_Addr_2.Text
                        txtShip_To_Addr_3.Text = txtBill_To_Addr_3.Text
                        txtShip_To_Addr_4.Text = txtBill_To_Addr_4.Text
                        txtShip_To_City.Text = txtBill_To_City.Text
                        txtShip_To_State.Text = txtBill_To_State.Text
                        txtShip_To_Zip.Text = txtBill_To_Zip.Text
                        txtShip_To_Country.Text = txtBill_To_Country.Text

                        m_oOrder.Ordhead.Ship_To_City = m_oOrder.Ordhead.Bill_To_City
                        m_oOrder.Ordhead.Ship_To_State = m_oOrder.Ordhead.Bill_To_State
                        m_oOrder.Ordhead.Ship_To_Zip = m_oOrder.Ordhead.Bill_To_Zip

                        txtCus_Alt_Adr_Cd.Text = "SAME"

                        If m_oOrder.Ordhead.IsARepOrder() Then
                            If txtShip_To_Name.Text = txtBill_To_Name.Text And txtShip_To_Addr_1.Text = txtBill_To_Addr_1.Text Then
                                If txtShip_Instruction_1.Text = String.Empty Then
                                    txtShip_Instruction_1.Text = "NO CHARGE FOR FREIGHT"
                                ElseIf txtShip_Instruction_2.Text = String.Empty And txtShip_Instruction_1.Text <> "NO CHARGE FOR FREIGHT" Then
                                    txtShip_Instruction_2.Text = "NO CHARGE FOR FREIGHT"
                                Else
                                    MsgBox("NO CHARGE FOR FREIGHT")
                                End If
                            End If
                        End If

                        blnSetTags = False

                    End If

                End If

            End If

            ' If no customer has been selected, then you can't select a ship to address.
            If piTabIndex >= lblBillToName.TabIndex And (txtCus_No.Text = "" Or txtCus_No.BackColor = Color.FromArgb(255, 192, 192)) Then

                txtCus_No.Focus()
                txtCus_No.BackColor = Color.FromArgb(255, 192, 192)
                ttToolTip.SetToolTip(txtCus_No, "Data must be entered here.")
                ForceFocus = True

                Exit Function

            End If

            ' If no customer has been selected, then you can't select a ship to address.
            If piTabIndex >= lblShipToName.TabIndex And (txtCus_Alt_Adr_Cd.Text = "" Or txtCus_Alt_Adr_Cd.BackColor = Color.FromArgb(255, 192, 192)) Then

                txtCus_Alt_Adr_Cd.Focus()
                txtCus_Alt_Adr_Cd.BackColor = Color.FromArgb(255, 192, 192)
                ttToolTip.SetToolTip(txtCus_No, "Data must be entered here.")
                ForceFocus = True

                Exit Function

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Function

    ' Reset every field after Customer name to String.Empty
    Private Sub ResetCustomerFields()

        Try

            txtOe_Po_No.Text = String.Empty
            txtSlspsn_No.Text = String.Empty
            txtShip_Link.Text = String.Empty
            txtJob_No.Text = String.Empty

            txtCus_Alt_Adr_Cd.Text = String.Empty
            txtMfg_Loc.Text = String.Empty
            txtShip_Via_Cd.Text = String.Empty
            txtAr_Terms_Cd.Text = String.Empty
            txtDiscount_Pct.Text = String.Empty
            txtProfit_Center.Text = String.Empty
            txtDept.Text = String.Empty
            txtShip_Instruction_1.Text = String.Empty
            txtShip_Instruction_2.Text = String.Empty

            txtBill_To_Name.Text = String.Empty
            txtBill_To_Addr_1.Text = String.Empty
            txtBill_To_Addr_2.Text = String.Empty
            txtBill_To_Addr_3.Text = String.Empty
            txtBill_To_Addr_4.Text = String.Empty
            txtBill_To_City.Text = String.Empty
            txtBill_To_State.Text = String.Empty
            txtBill_To_Zip.Text = String.Empty
            txtBill_To_Country.Text = String.Empty
            lblBillToCountryDesc.Text = String.Empty

            txtShip_To_Name.Text = String.Empty
            txtShip_To_Addr_1.Text = String.Empty
            txtShip_To_Addr_2.Text = String.Empty
            txtShip_To_Addr_3.Text = String.Empty
            txtShip_To_Addr_4.Text = String.Empty
            txtShip_To_City.Text = String.Empty
            txtShip_To_State.Text = String.Empty
            txtShip_To_Zip.Text = String.Empty
            txtShip_To_Country.Text = String.Empty
            lblShipToCountryDesc.Text = String.Empty

            txtEnd_User.Text = String.Empty
            txtUser_Def_Fld_1.Text = String.Empty
            cboUser_Def_Fld_2.SelectedText = String.Empty
            '++ID 06.17.2020 extra_9
            cbo_Extra_9.SelectedText = String.Empty

            txtUser_Def_Fld_3.Text = String.Empty
            txtUser_Def_Fld_4.Text = String.Empty
            txtUser_Def_Fld_5.Text = String.Empty

            txtProg_Spector_Cd.Text = String.Empty
            '++ID 8.15.2018
            txtCustGroup.Text = String.Empty
            lblVIP.Visible = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

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

        Try

            m_oOrder.Ordhead.Cus_No = txtCus_No.Text
            m_oOrder.Ordhead.CheckExistingNumber(txtOrd_No)

            ' Before et get customer default values, we must get the order date for the rate.
            If IsDate(txtOrd_Dt.Text) Then
                m_oOrder.Ordhead.Ord_Dt = txtOrd_Dt.Text
            End If

            If m_oOrder.Ordhead.GetCustomerDefaultValues() Then

                Select Case m_oOrder.Ordhead.Customer.cmp_status
                    Case "A", "S" ' Passive and active, ok to proceed
                        m_oOrder.Ordhead.Hold_Fg = ""
                    Case "B" ' Blocked, show message. On message if not OK, restart order.
                        Dim mbResult As New MsgBoxResult
                        mbResult = MsgBox(New OEExceptionMessage(OEError.Cust_On_Credit_Hold).Message, MsgBoxStyle.YesNo, "Error.")
                        If mbResult = MsgBoxResult.No Then
                            ' Cancel the order
                            txtCus_No.Text = String.Empty
                            txtCus_No.Focus()
                            Exit Sub
                        Else
                            m_oOrder.Ordhead.Hold_Fg = "H"
                        End If
                    Case Else '"E" - treat as inactive
                        MsgBox("Customer is inactive.", MsgBoxStyle.OkOnly, "Error.")
                        txtCus_No.Text = String.Empty
                        txtCus_No.Focus()
                        Exit Sub
                End Select

                'Specific instruction popup (if applicable)
                If m_oOrder.Ordhead.Customer.Specific_Instructions <> "" Then
                    MsgBox(New OEExceptionMessage(OEError.Specific_Instructions).Message & ":" & vbCrLf & vbCrLf & m_oOrder.Ordhead.Customer.Specific_Instructions, MsgBoxStyle.OkOnly)
                End If

                txtCus_Alt_Adr_Cd.Text = String.Empty

                txtOe_Po_No.Text = String.Empty
                txtSlspsn_No.Text = m_oOrder.Ordhead.Slspsn_No.ToString
                txtShip_Link.Text = String.Empty
                txtJob_No.Text = String.Empty

                txtMfg_Loc.Text = m_oOrder.Ordhead.Mfg_Loc
                txtShip_Via_Cd.Text = m_oOrder.Ordhead.Ship_Via_Cd
                txtAr_Terms_Cd.Text = m_oOrder.Ordhead.Ar_Terms_Cd
                txtDiscount_Pct.Text = 0 '++ID 7.5.2018 put in comment is dangerouse for use this data from Macola  m_oOrder.Ordhead.Discount_Pct
                txtProfit_Center.Text = m_oOrder.Ordhead.Profit_Center

                txtDept.Text = String.Empty

                txtShip_Instruction_1.Text = String.Empty
                txtShip_Instruction_2.Text = m_oOrder.Ordhead.Customer.Comment1 ' String.Empty

                txtBill_To_Name.Text = m_oOrder.Ordhead.Bill_To_Name
                txtBill_To_Addr_1.Text = m_oOrder.Ordhead.Bill_To_Addr_1
                txtBill_To_Addr_2.Text = m_oOrder.Ordhead.Bill_To_Addr_2
                txtBill_To_Addr_3.Text = m_oOrder.Ordhead.Bill_To_Addr_3
                txtBill_To_Addr_4.Text = m_oOrder.Ordhead.Bill_To_Addr_4
                txtBill_To_City.Text = m_oOrder.Ordhead.Bill_To_City
                txtBill_To_State.Text = m_oOrder.Ordhead.Bill_To_State
                txtBill_To_Zip.Text = m_oOrder.Ordhead.Bill_To_Zip
                txtBill_To_Country.Text = m_oOrder.Ordhead.Bill_To_Country

                'txtBill_To_CountryDesc.Text = rsCustomer.Fields("").Value 'will populate on country change

                'm_oOrder.Ordhead.Curr_Cd = IIf(System.DBNull.Value.Equals(rsCustomer.Fields("Curr_Cd").Value), "", rsCustomer.Fields("Curr_Cd").Value)

                txtShip_To_Name.Text = String.Empty
                txtShip_To_Addr_1.Text = String.Empty
                txtShip_To_Addr_2.Text = String.Empty
                txtShip_To_Addr_3.Text = String.Empty
                txtShip_To_Addr_4.Text = String.Empty
                txtShip_To_City.Text = String.Empty
                txtShip_To_State.Text = String.Empty
                txtShip_To_Zip.Text = String.Empty
                txtShip_To_Country.Text = String.Empty
                lblShipToCountryDesc.Text = String.Empty

                cboUser_Def_Fld_2.SelectedText = IIf(m_oOrder.Ordhead.User_Def_Fld_2 Is Nothing, "", Trim(m_oOrder.Ordhead.User_Def_Fld_2))
                cboUser_Def_Fld_2.Text = IIf(m_oOrder.Ordhead.User_Def_Fld_2 Is Nothing, "", Trim(m_oOrder.Ordhead.User_Def_Fld_2))

                '++ID 06.17.2020 added new field only for oei_ordheader
                cbo_Extra_9.SelectedText = IIf(m_oOrder.Ordhead.Extra_9 Is Nothing, "", Trim(m_oOrder.Ordhead.Extra_9))
                cbo_Extra_9.Text = IIf(m_oOrder.Ordhead.Extra_9 Is Nothing, "", Trim(m_oOrder.Ordhead.Extra_9))
                '----------------

                txtUser_Def_Fld_3.Text = m_oOrder.Ordhead.User_Def_Fld_3
                txtContactEmail.Text = m_oOrder.Ordhead.Cus_Email_Address

                '++ID 12.12.2024 show customer EIN
                txtShip_Instruction_1.Text &= "EIN:" & m_oOrder.Ordhead.Customer.TextField7 & ". "

                'moved before  function save and not after
                txtCustGroup.Text = m_oOrder.Ordhead.Customer.textfield3
                ttToolTip.SetToolTip(txtCustGroup, "Customer Group Name")

                Save()

                '++ID 8.14.2018 
                'txtCustGroup.Text = m_oOrder.Ordhead.Customer.textfield3
                'ttToolTip.SetToolTip(txtCustGroup, "Customer Group Name")


                If m_oOrder.Ordhead.Customer.TextField6 <> String.Empty Then
                    lblVIP.Visible = True
                Else
                    lblVIP.Visible = False
                End If


                'Else
                'Call ResetCustomerFields()
                'End If

                'txtCus_No.Enabled = (Trim(txtCus_No.Text.Length) = "") ''Not (rsCustomer.RecordCount <> 0)
                'cmdCustomerSearch.Enabled = txtCus_No.Enabled

                txtOrd_No.Enabled = (Trim(txtCus_No.Text.Length) = "")
                cmdOrderSearch.Enabled = txtOrd_No.Enabled
                cboOrd_Type.Enabled = txtOrd_No.Enabled
                cmdCopyOrderFromHistory.Enabled = txtOrd_No.Enabled
                chkRecover.Enabled = (Trim(txtCus_No.Text) = "")
                chkExactRepeat.Enabled = chkRecover.Enabled ' (Trim(txtCus_No.Text) = "")
                'rsCustomer.Close()

            Else
                txtCus_No.Text = String.Empty
                txtCus_No.Focus()
                Exit Sub
            End If

            blnSetTags = False

            Dim dt As DataTable
            Dim db As New cDBA
            Dim strSql As String = "SELECT ID, Cmt AS Comments, Cus_Sort FROM Exact_Custom_OeFlash_CustomerComments WHERE Cus_no = '" & m_oOrder.Ordhead.Cus_No & "' "
            dt = db.DataTable(strSql)

            If ofrmOEFlash Is Nothing Then
                If dt.Rows.Count <> 0 Then ' Then
                    cmdOEFlash.PerformClick()
                End If
            ElseIf dt.Rows.Count <> 0 Or ofrmOEFlash.Visible Then
                cmdOEFlash.PerformClick()
            End If

            dt.Dispose()

            m_oOrder.CreditInfo = New cCreditInfo(m_oOrder.Ordhead.Cus_No)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function GetProfit_Center(ByVal pstrCurr_Cd As String) As String

        GetProfit_Center = ""

        'Dim rsPC As New ADODB.Recordset

        Dim strsql As String = _
        "SELECT ISNULL(Sls_Sb_No, '') AS Sls_Sb_No " & _
        "FROM   Artypfil_Sql WITH (NoLock) " & _
        "WHERE  CUS_TYPE_CD = '" & pstrCurr_Cd & "'"

        Dim dt As DataTable
        Dim db As New cDBA

        dt = db.DataTable(strsql)
        If dt.Rows.Count <> 0 Then
            GetProfit_Center = dt.Rows(0).Item("sls_sb_no")
        End If

        'rsPC.Open(strsql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'If rsPC.RecordCount <> 0 Then
        '    GetProfit_Center = rsPC.Fields("sls_sb_no").Value
        'End If

        'rsPC.Close()

    End Function

    'Private Sub SetAlternateAddress(ByVal pstrCus_No As String, ByVal pstrCus_Alt_Adr_Cd As String)

    '    Try
    '        Dim rsAltAdr As New ADODB.Recordset
    '        Dim strSql As String = "" & _
    '        "SELECT     TOP 1 WITH TIES * " & _
    '        "FROM       ARALTADR_SQL WITH (Nolock)  " & _
    '        "WHERE     (RTRIM(ARALTADR_SQL.cus_no) LIKE '" & RTrim(LTrim(pstrCus_No)) & "%' OR RTRIM(ARALTADR_SQL.cus_no) LIKE '% " & RTrim(LTrim(pstrCus_No)) & "%') AND " & _
    '        "		    CUS_ALT_ADR_CD = '" & SqlCompliantString(pstrCus_Alt_Adr_Cd) & "' " & _
    '        "ORDER BY   ARALTADR_SQL.cus_alt_adr_cd "

    '        rsAltAdr.Open(strSql, gConn.ConnectionString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '        If rsAltAdr.RecordCount <> 0 Then
    '            'CUS_ALT_ADR_CD -> DS LE CODE 
    '            'ALTERNATE_ADDRESS -> DS LE NOM ADRESSE
    '            '            ADDR_1()
    '            '            ADDR_2()
    '            '            ADDR_3()
    '            'CITY & ", " & STATE & " " & ZIP DANS ADDR_4
    '            '            COUNTRY()
    '            txtShip_To_Name.Text = IIf(rsAltAdr.Fields("user_def_fld_1").Equals(DBNull.Value), "", rsAltAdr.Fields("user_def_fld_1").Value)
    '            txtShip_To_Addr_1.Text = IIf(rsAltAdr.Fields("ADDR_1").Equals(DBNull.Value), "", rsAltAdr.Fields("ADDR_1").Value)
    '            txtShip_To_Addr_2.Text = IIf(rsAltAdr.Fields("ADDR_2").Equals(DBNull.Value), "", rsAltAdr.Fields("ADDR_2").Value)
    '            txtShip_To_Addr_3.Text = IIf(rsAltAdr.Fields("ADDR_3").Equals(DBNull.Value), "", rsAltAdr.Fields("ADDR_3").Value)
    '            txtShip_To_Addr_4.Text = IIf(rsAltAdr.Fields("CITY").Equals(DBNull.Value), "", rsAltAdr.Fields("CITY").Value & ", ")
    '            txtShip_To_Addr_4.Text = txtShip_To_Addr_4.Text & IIf(rsAltAdr.Fields("STATE").Equals(DBNull.Value), "", rsAltAdr.Fields("STATE").Value & " ")
    '            txtShip_To_Addr_4.Text = txtShip_To_Addr_4.Text & IIf(rsAltAdr.Fields("ZIP").Equals(DBNull.Value), "", rsAltAdr.Fields("ZIP").Value)
    '            txtShip_To_Country.Text = IIf(rsAltAdr.Fields("COUNTRY").Equals(DBNull.Value), "", rsAltAdr.Fields("COUNTRY").Value)

    '            m_oOrder.Ordhead.Ship_To_City = IIf(rsAltAdr.Fields("CITY").Equals(DBNull.Value), "", rsAltAdr.Fields("CITY").Value)
    '            m_oOrder.Ordhead.Ship_To_State = IIf(rsAltAdr.Fields("STATE").Equals(DBNull.Value), "", rsAltAdr.Fields("STATE").Value)
    '            m_oOrder.Ordhead.Ship_To_Zip = IIf(rsAltAdr.Fields("ZIP").Equals(DBNull.Value), "", rsAltAdr.Fields("ZIP").Value)

    '        End If
    '        rsAltAdr.Close()
    '        rsAltAdr = Nothing

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
    '    End Try

    'End Sub

    Private Sub cmdCopyOrderFromHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyOrderFromHistory.Click

        Try
            m_oOrder = New cOrder(txtOrd_No.Text, OrderSourceEnum.Macola)
            Call Fill() ' m_oOrder)
            txtOrd_No.Text = "NEW"
            m_oOrder.FromHistory = True
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub LoadOrderFromOEI(ByVal pstrOrd_No As String)

        If pstrOrd_No Is Nothing Then Exit Sub

        Try
            m_oOrder = New cOrder(pstrOrd_No, OrderSourceEnum.OEInterface)
            Call Fill() ' m_oOrder)
            m_oOrder.FromHistory = False

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub CreateExactRepeatOrder(ByVal pstrOrd_No As String)

        Try
            Dim dtHeader As New DataTable
            Dim dtLines As New DataTable
            Dim db As New cDBA
            Dim dtRow As DataRow = Nothing

            ' First part - create order header

            m_oOrder = New cOrder
            m_oOrder.Ordhead.OEI_Ord_No = g_User.NextOrderNumber()
            m_oOrder.Ordhead.Ord_Dt = Now.Date
            m_oOrder.Ordhead.Shipping_Dt = Now.Date
            m_oOrder.Ordhead.Extra_6 = Trim(pstrOrd_No).PadLeft(8)

            Dim strSql As String = ""
            strSql = _
            "SELECT * " & _
            "FROM   OEHDRHST_SQL WITH (Nolock) " & _
            "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

            dtHeader = db.DataTable(strSql)

            If dtHeader.Rows.Count = 0 Then

                strSql = _
                "SELECT * " & _
                "FROM   OEORDHDR_SQL WITH (Nolock) " & _
                "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

                dtHeader = db.DataTable(strSql)

            Else
                m_oOrder.FromHistory = True
            End If

            If dtHeader.Rows.Count <> 0 Then

                dtRow = dtHeader.Rows(0)

                With m_oOrder.Ordhead

                    .Ord_Type = Trim(dtRow.Item("Ord_Type").ToString)
                    .Cus_No = Trim(dtRow.Item("Cus_No").ToString)
                    Call .GetCustomerDefaultValues()

                End With

                Select Case m_oOrder.Ordhead.Customer.cmp_status
                    Case "A", "S" ' Passive and active, ok to proceed
                        m_oOrder.Ordhead.Hold_Fg = ""
                    Case "B" ' Blocked, show message. On message if not OK, restart order.
                        Dim mbResult As New MsgBoxResult
                        mbResult = MsgBox(New OEExceptionMessage(OEError.Cust_On_Credit_Hold).Message, MsgBoxStyle.YesNo, "Warning.")
                        If mbResult = MsgBoxResult.No Then
                            Throw New OEException(OEError.Cust_On_Credit_Hold, True, False)
                            ' Cancel the order
                            'txtCus_No.Text = String.Empty
                            'txtCus_No.Focus()
                            'Exit Sub
                        Else
                            m_oOrder.Ordhead.Hold_Fg = "H"
                        End If
                    Case Else '"E" - treat as inactive
                        Throw New OEException(OEError.Cust_Is_Inactive, True, True)
                        'MsgBox("Customer is inactive.", MsgBoxStyle.OkOnly, "Error.")
                        'txtCus_No.Text = String.Empty
                        'txtCus_No.Focus()
                        'Exit Sub
                End Select
            End If

            'Specific instruction popup (if applicable)
            If m_oOrder.Ordhead.Customer.Specific_Instructions <> "" Then
                MsgBox(New OEExceptionMessage(OEError.Specific_Instructions).Message & ":" & vbCrLf & vbCrLf & m_oOrder.Ordhead.Customer.Specific_Instructions, MsgBoxStyle.OkOnly)
            End If

            If dtHeader.Rows.Count <> 0 Then

                'Dim dtRow As DataRow = dtHeader.Rows(0)
                With m_oOrder.Ordhead
                    '.Ord_Type = Trim(dtRow.Item("Ord_Type").ToString)
                    '.Cus_No = Trim(dtRow.Item("Cus_No").ToString)
                    'Call .GetCustomerDefaultValues()
                    '.Cus_Alt_Adr_Cd = Trim(dtRow.Item("Cus_Alt_Adr_Cd").ToString)
                    '.Slspsn_No = Trim(dtRow.Item("Slspsn_No").ToString)
                    '.Slspsn_Pct_Comm = dtRow.Item("Slspsn_Pct_Comm")


                    .User_Def_Fld_4 = Trim(dtRow.Item("User_Def_Fld_4").ToString)
                    .Email_Address = Trim(dtRow.Item("Email_Address").ToString)
                    '.Mfg_Loc = Trim(dtRow.Item("Mfg_Loc").ToString)
                    '.Ship_Via_Cd = Trim(dtRow.Item("Ship_Via_Cd").ToString)
                    '.Ar_Terms_Cd = Trim(dtRow.Item("Ar_Terms_Cd").ToString)
                    '.Profit_Center = Trim(dtRow.Item("Profit_Center").ToString)
                    '.Dept = Trim(dtRow.Item("Dept").ToString)
                    '.User_Def_Fld_2 = Trim(dtRow.Item("User_Def_Fld_2").ToString)
                    '.Discount_Pct = Trim(dtRow.Item("Discount_Pct").ToString)
                    '.End_User = Trim(dtRow.Item("End_User").ToString)
                    .User_Def_Fld_1 = Trim(dtRow.Item("User_Def_Fld_1").ToString)
                    '.User_Def_Fld_3 = Trim(dtRow.Item("User_Def_Fld_3").ToString)
                    .User_Def_Fld_5 = Trim(dtRow.Item("User_Def_Fld_5").ToString)
                    .Ship_Instruction_1 = Trim(dtRow.Item("Ship_Instruction_1").ToString)
                    .Ship_Instruction_2 = Trim(dtRow.Item("Ship_Instruction_2").ToString)
                    '.Bill_To_Name = Trim(dtRow.Item("Bill_To_Name").ToString)
                    '.Bill_To_Addr_1 = Trim(dtRow.Item("Bill_To_Addr_1").ToString)
                    '.Bill_To_Addr_2 = Trim(dtRow.Item("Bill_To_Addr_2").ToString)
                    '.Bill_To_Addr_3 = Trim(dtRow.Item("Bill_To_Addr_3").ToString)
                    '.Bill_To_Addr_4 = Trim(dtRow.Item("Bill_To_Addr_4").ToString)
                    '.Bill_To_City = Trim(dtRow.Item("Bill_To_City").ToString)
                    '.Bill_To_State = Trim(dtRow.Item("Bill_To_State").ToString)
                    '.Bill_To_Zip = Trim(dtRow.Item("Bill_To_Zip").ToString)
                    '.Bill_To_Country = Trim(dtRow.Item("Bill_To_Country").ToString)
                    .Ship_To_Name = Trim(dtRow.Item("Ship_To_Name").ToString)
                    .Ship_To_Addr_1 = Trim(dtRow.Item("Ship_To_Addr_1").ToString)
                    .Ship_To_Addr_2 = Trim(dtRow.Item("Ship_To_Addr_2").ToString)
                    .Ship_To_Addr_3 = Trim(dtRow.Item("Ship_To_Addr_3").ToString)
                    .Ship_To_Addr_4 = Trim(dtRow.Item("Ship_To_Addr_4").ToString)
                    .Ship_To_City = Trim(dtRow.Item("Ship_To_City").ToString)
                    .Ship_To_State = Trim(dtRow.Item("Ship_To_State").ToString)
                    .Ship_To_Zip = Trim(dtRow.Item("Ship_To_Zip").ToString)
                    .Ship_To_Country = Trim(dtRow.Item("Ship_To_Country").ToString)

                    '.Curr_Cd = Trim(dtRow.Item("Curr_Cd").ToString)
                    .Orig_Trx_Rt = dtRow.Item("Curr_Trx_Rt")
                    '.Curr_Trx_Rt = dtRow.Item("Curr_Trx_Rt")
                    '.Form_No = dtRow.Item("Form_No")
                    '.Deter_Rate_By = dtRow.Item("Deter_Rate_By")
                    '.Tax_Fg = Trim(dtRow.Item("Tax_Fg").ToString)
                    .Contact_1 = Trim(dtRow.Item("Contact_1").ToString)
                    .Phone_Number = Trim(dtRow.Item("Phone_Number").ToString)
                    .Fax_Number = Trim(dtRow.Item("Fax_Number").ToString)
                    .Email_Address = Trim(dtRow.Item("Email_Address").ToString)
                    .Cus_Email_Address = Trim(dtRow.Item("Email_Address").ToString)
                    ' .in_hands_date = Trim(dtRow.Item("In_Hands_Date").ToString)
                    ' .Cus_Email_Address = Trim(dtRow.Item("In_Hands_Date").ToString)

                    .Save()

                End With

            End If

            ' second part - create order lines

            If m_oOrder.FromHistory Then

                strSql = _
                "SELECT * " & _
                "FROM   OELINHST_SQL WITH (Nolock) " & _
                "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

            Else

                strSql = _
                "SELECT * " & _
                "FROM   OEORDLIN_SQL WITH (Nolock) " & _
                "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

            End If

            dtLines = db.DataTable(strSql)

            If dtLines.Rows.Count <> 0 And dtHeader.Rows.Count <> 0 Then
                For Each dtRow In dtLines.Rows

                    Dim oOrdline As New cOrdLine(m_oOrder.Ordhead.Ord_GUID)
                    oOrdline.Source = OEI_SourceEnum.Macola_ExactRepeat
                    oOrdline.Line_Seq_No = dtRow.Item("Line_Seq_No")
                    oOrdline.OEOrdLin_ID = dtRow.Item("Line_Seq_No")
                    oOrdline.Item_No = dtRow.Item("Item_No").ToString
                    oOrdline.Loc = dtRow.Item("Loc").ToString

                    ' Qty to ship is set to qty ordered, because of overpicks.
                    Try
                        oOrdline.Qty_Ordered = dtRow.Item("Qty_Ordered")
                    Catch oe_er As OEException
                        'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                    End Try

                    Try
                        If oOrdline.Qty_Prev_Bkord > 0 Then
                            oOrdline.Qty_To_Ship = oOrdline.Qty_Ordered - oOrdline.Qty_Prev_Bkord
                            'oOrdline.Calculate_ETA()
                        Else
                            oOrdline.Qty_To_Ship = dtRow.Item("Qty_Ordered")
                        End If

                        'MAYBE WE NEED TI SKIP THIS PART ??
                        'oOrdline.Qty_To_Ship_Changed(oOrdline.Item_No, oOrdline.Loc, False)
                        'If oOrdline.Qty_Prev_Bkord > 0 Then
                        '    'oOrdline.Calculate_ETA()
                        'End If

                    Catch oe_er As OEException
                        'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
                    End Try

                    ' Do not enter unit price, let calculate new price on date
                    ' oOrdline.Unit_Price = dtRow.Item("Unit_Price").ToString
                    oOrdline.Discount_Pct = dtRow.Item("Discount_Pct")
                    oOrdline.Request_Dt = Now.Date
                    oOrdline.Promise_Dt = Now.Date
                    oOrdline.Req_Ship_Dt = Now.Date
                    oOrdline.Item_Desc_1 = dtRow.Item("Item_Desc_1").ToString
                    oOrdline.Item_Desc_2 = dtRow.Item("Item_Desc_2").ToString
                    oOrdline.Extra_1 = dtRow.Item("Extra_1").ToString

                    oOrdline.Source = OEI_SourceEnum.frmOrder
                    oOrdline.Save()

                Next dtRow

            End If

            ' Third part - create traveler. We cannot create in lines
            ' because of kit items.
            strSql = "" & _
            "SELECT 	L.*, ISNULL(B.Seq_No, 0) AS Kit_No " & _
            "FROM 		OEI_OrdLin L " & _
            "LEFT JOIN	OEI_OrdBld B WITH (NoLock) ON L.Ord_Guid = B.Ord_Guid AND L.Item_Guid = B.Comp_Item_Guid " & _
            "WHERE      L.Ord_Guid = '" & Trim(m_oOrder.Ordhead.Ord_GUID) & "' " & _
            "ORDER BY   L.Line_No, B.Seq_No "

            dtLines = db.DataTable(strSql)

            If dtLines.Rows.Count <> 0 Then

                For Each dtRow In dtLines.Rows

                    'If m_oOrder.OrdLines.Contains(dtRow.Item("Item_Guid").ToString) Then

                    Dim dtItem As DataTable

                    strSql = "" & _
                    "SELECT 	ISNULL(RouteID, 0) AS RouteID " & _
                    "FROM 		Exact_Traveler_Header WITH (Nolock) " & _
                    "WHERE		Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND " & _
                    "   		Line_No = " & dtRow.Item("OEOrdLin_ID") & " AND " & _
                    "           Component_Seq_No = " & dtRow.Item("Kit_No") & " " & _
                    "ORDER BY 	Ord_No, Line_No, Component_Seq_No "

                    dtItem = db.DataTable(strSql)
                    If dtItem.Rows.Count <> 0 Then

                        strSql = _
                        "UPDATE OEI_OrdLin " & _
                        "SET    RouteID = " & dtItem.Rows(0).Item("RouteID") & " " & _
                        "WHERE  Ord_Guid = '" & dtRow.Item("Ord_Guid") & "' AND " & _
                        "       Item_Guid = '" & dtRow.Item("Item_Guid") & "'"

                        db.Execute(strSql)
                        'db.DBCommand.CommandText = strSql
                        'db.DBCommand.ExecuteNonQuery()
                        'db.DBCommand.CommandText = ""

                        'Dim oImprint As New cImprint(Trim(m_oOrder.Ordhead.Ord_GUID), Trim(dtItem.Rows(0).Item("Item_Guid").ToString))
                        ''oImprint.Ord_Type = dtRow.Item("Ord_Type").ToString
                        ''oImprint.OE_Line_No = dtRow.Item("OE_Line_No").ToString
                        ''oImprint.Comp_Seq_No = dtRow.Item("Comp_Seq_No").ToString
                        'oImprint.Item_No = dtRow.Item("Item_No").ToString
                        ''oImprint.Par_Item_No = dtRow.Item("Par_Item_No").ToString
                        'oImprint.Imprint = dtRow.Item("Extra_1").ToString
                        'oImprint.Location = dtRow.Item("Extra_2").ToString
                        'oImprint.Color = dtRow.Item("Extra_3").ToString
                        'If IsNumeric(dtRow.Item("Extra_4").ToString) Then
                        '    oImprint.Num = dtRow.Item("Extra_4").ToString
                        'Else
                        '    oImprint.Num = 0
                        'End If
                        'oImprint.Packaging = dtRow.Item("Extra_5").ToString
                        ''oImprint.Extra_6 = dtRow.Item("Extra_6").ToString
                        ''oImprint.Extra_7 = dtRow.Item("Extra_7").ToString
                        'oImprint.Refill = dtRow.Item("Extra_8").ToString
                        'oImprint.Laser_Setup = dtRow.Item("Extra_9").ToString
                        'oImprint.Industry = dtRow.Item("Industry").ToString
                        'oImprint.Comments = dtRow.Item("Comment").ToString
                        'oImprint.Special_Comments = "REPEAT ORDER # " & Trim(pstrOrd_No) ' dtRow.Item("Comment2").ToString
                        'oImprint.Save()

                    End If
                    'End If

                Next dtRow

            End If


            ' fourt part - create order imprint data. We cannot create in lines
            ' because of kit items.

            'strSql = "" & _
            '"SELECT 	EF.*, ISNULL(L.Line_No, 0) AS OEOrdLin_ID, ISNULL(EF.Comp_Seq_No, 0) AS Kit_No " & _
            '"FROM 		OELinHst_Sql L WITH (Nolock) " & _
            '"LEFT JOIN	Exact_Traveler_Extra_Fields EF WITH (Nolock) ON L.Ord_No = EF.Ord_No AND L.Line_Seq_No = EF.OE_Line_No " & _
            '"LEFT JOIN	IMOrdHst_Sql B WITH (Nolock) ON EF.Ord_No = B.Ord_No AND EF.OE_Line_No = B.Line_No AND EF.Comp_Seq_No = B.Seq_No " & _
            '"WHERE		L.Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' " & _
            '"ORDER BY 	L.Ord_No, L.Line_Seq_No, EF.Comp_Seq_No "

            ' No need to redo, will use the same one as previous query

            'dtLines = db.DataTable(strSql)

            If dtLines.Rows.Count <> 0 Then

                For Each dtRow In dtLines.Rows

                    Dim dtItem As DataTable

                    'strSql = "" & _
                    '"SELECT 	L.* " & _
                    '"FROM 		OEI_OrdLin L WITH (NoLock) " & _
                    '"LEFT JOIN	OEI_OrdBld B WITH (NoLock) ON L.Ord_Guid = B.Ord_Guid AND L.Item_Guid = B.Comp_Item_Guid " & _
                    '"WHERE      L.Ord_Guid = '" & Trim(m_oOrder.Ordhead.Ord_GUID) & "' AND " & _
                    '"           L.Line_Seq_No = " & dtRow.Item("OEOrdLin_ID") & " " & _
                    'IIf(dtRow.Item("Kit_No") = 0, "", " AND B.Seq_No = " & dtRow.Item("Kit_No") & " ") & _
                    '"ORDER BY   L.Line_No, B.Seq_No "

                    strSql = "" & _
                    "SELECT 	*, CASE WHEN ISNULL(Num_Impr_1, 0) = 0 THEN CONVERT(INT, ISNULL(Extra_4, '0')) ELSE ISNULL(Num_Impr_1, 0) END AS Comp_Num_Impr " & _
                    "FROM 		Exact_Traveler_Extra_Fields WITH (Nolock) " & _
                    "WHERE		Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND " & _
                    "   		OE_Line_No = " & dtRow.Item("OEOrdLin_ID") & " AND " & _
                    "           Comp_Seq_No = " & dtRow.Item("Kit_No") & " " & _
                    "ORDER BY 	Ord_No, OE_Line_No, Comp_Seq_No "

                    dtItem = db.DataTable(strSql)
                    If dtItem.Rows.Count <> 0 Then

                        Dim oImprint As New cImprint(Trim(m_oOrder.Ordhead.Ord_GUID), Trim(dtRow.Item("Item_Guid").ToString))
                        'oImprint.Ord_Type = dtRow.Item("Ord_Type").ToString
                        'oImprint.OE_Line_No = dtRow.Item("OE_Line_No").ToString
                        'oImprint.Comp_Seq_No = dtRow.Item("Comp_Seq_No").ToString
                        oImprint.SaveToDB = False
                        With dtItem.Rows(0)

                            oImprint.Item_No = .Item("Item_No").ToString
                            'oImprint.Par_Item_No = dtRow.Item("Par_Item_No").ToString
                            oImprint.Imprint = .Item("Extra_1").ToString
                            oImprint.Location = .Item("Extra_2").ToString
                            oImprint.Color = .Item("Extra_3").ToString
                            If IsNumeric(.Item("Comp_Num_Impr").ToString) Then
                                oImprint.Num_Impr_1 = .Item("Comp_Num_Impr")
                            Else
                                oImprint.Num_Impr_1 = 0
                            End If
                            If IsNumeric(.Item("Num_Impr_2").ToString) Then
                                oImprint.Num_Impr_2 = .Item("Num_Impr_2")
                            Else
                                oImprint.Num_Impr_2 = 0
                            End If
                            If IsNumeric(.Item("Num_Impr_3").ToString) Then
                                oImprint.Num_Impr_3 = .Item("Num_Impr_3")
                            Else
                                oImprint.Num_Impr_3 = 0
                            End If
                            oImprint.Packaging = .Item("Extra_5").ToString
                            'oImprint.Extra_6 = dtRow.Item("Extra_6").ToString
                            'oImprint.Extra_7 = dtRow.Item("Extra_7").ToString
                            oImprint.Refill = .Item("Extra_8").ToString
                            oImprint.Laser_Setup = .Item("Extra_9").ToString
                            oImprint.Industry = .Item("Industry").ToString
                            oImprint.Comments = .Item("Comment").ToString
                            oImprint.Repeat_From_Ord_No = .Item("Ord_No").ToString
                            oImprint.Repeat_From_ID = .Item("ID")
                            oImprint.Special_Comments = "REPEAT ORDER # " & Trim(pstrOrd_No)
                            oImprint.Printer_Comment = .Item("Printer_Comment").ToString
                            oImprint.Packer_Comment = .Item("Packer_Comment").ToString
                            oImprint.Printer_Instructions = .Item("Printer_Instructions").ToString
                            oImprint.Packer_Instructions = .Item("Packer_Instructions").ToString
                            oImprint.End_User = .Item("End_User").ToString

                        End With

                        g_oOrdline.Set_Spec_Instructions(oImprint.Item_No, oImprint)
                        oImprint.SaveToDB = True

                        If Not oImprint.IsEmpty() Then oImprint.Save()

                    Else

                        Dim oImprint As New cImprint(Trim(m_oOrder.Ordhead.Ord_GUID), Trim(dtRow.Item("Item_Guid").ToString))

                        oImprint.SaveToDB = False
                        oImprint.Item_No = dtRow.Item("Item_No").ToString
                        oImprint.Num_Impr_1 = 1

                        g_oOrdline.Set_Spec_Instructions(oImprint.Item_No, oImprint)
                        oImprint.SaveToDB = True

                        If Not oImprint.IsEmpty() Then oImprint.Save()

                    End If

                Next dtRow

            End If

            strSql = "DELETE FROM OEI_Order_Contacts WHERE ORD_GUID = '" & m_oOrder.Ordhead.Ord_GUID & "' "
            db.Execute(strSql)

            strSql = _
            "INSERT INTO OEI_Order_Contacts (Ord_Guid, DefContact, DefMethod, ContactType) " & _
            "SELECT 	'" & m_oOrder.Ordhead.Ord_GUID & "', DefContact, DefMethod, ContactType  " & _
            "FROM 		Exact_Traveler_Order_Customer_Contacts_New " & _
            "WHERE 		OrderNo = '" & Trim(pstrOrd_No) & "'"
            db.Execute(strSql)

            m_oOrder.FromHistory = False

        Catch oe_er As OEException
            Throw oe_er ' New OEException(oe_er.Message, oe_er)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub CreateEDIOrderFromXML()

        'Try
        '    Dim oXmlFile As Xml.XmlDocument

        '    Dim dtHeader As New DataTable
        '    Dim dtLines As New DataTable
        '    Dim db As New cDBA
        '    Dim dtRow As DataRow = Nothing

        '    ' First part - create order header

        '    m_oOrder = New cOrder

        '    m_oOrder.Ordhead.OEI_Ord_No = g_User.NextOrderNumber()


        '    m_oOrder.Ordhead.Ord_Dt = Now.Date
        '    m_oOrder.Ordhead.Shipping_Dt = Now.Date
        '    m_oOrder.Ordhead.Extra_6 = Trim(pstrOrd_No).PadLeft(8)

        '    Dim strSql As String = ""
        '    strSql = _
        '    "SELECT * " & _
        '    "FROM   OEHDRHST_SQL WITH (Nolock) " & _
        '    "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

        '    dtHeader = db.DataTable(strSql)

        '    If dtHeader.Rows.Count = 0 Then

        '        strSql = _
        '        "SELECT * " & _
        '        "FROM   OEORDHDR_SQL WITH (Nolock) " & _
        '        "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

        '        dtHeader = db.DataTable(strSql)

        '    Else
        '        m_oOrder.FromHistory = True
        '    End If

        '    If dtHeader.Rows.Count <> 0 Then

        '        dtRow = dtHeader.Rows(0)

        '        With m_oOrder.Ordhead

        '            .Ord_Type = Trim(dtRow.Item("Ord_Type").ToString)
        '            .Cus_No = Trim(dtRow.Item("Cus_No").ToString)
        '            Call .GetCustomerDefaultValues()

        '        End With

        '        Select Case m_oOrder.Ordhead.Customer.cmp_status
        '            Case "A", "S" ' Passive and active, ok to proceed
        '            Case "B" ' Blocked, show message. On message if not OK, restart order.
        '                Dim mbResult As New MsgBoxResult
        '                mbResult = MsgBox(New OEExceptionMessage(OEError.Cust_On_Credit_Hold).Message, MsgBoxStyle.YesNo, "Warning.")
        '                If mbResult = MsgBoxResult.No Then
        '                    Throw New OEException(OEError.Cust_On_Credit_Hold, True, False)
        '                    ' Cancel the order
        '                    'txtCus_No.Text = String.Empty
        '                    'txtCus_No.Focus()
        '                    'Exit Sub
        '                Else
        '                    m_oOrder.Ordhead.Hold_Fg = "H"
        '                End If
        '            Case Else '"E" - treat as inactive
        '                Throw New OEException(OEError.Cust_Is_Inactive, True, True)
        '                'MsgBox("Customer is inactive.", MsgBoxStyle.OkOnly, "Error.")
        '                'txtCus_No.Text = String.Empty
        '                'txtCus_No.Focus()
        '                'Exit Sub
        '        End Select
        '    End If

        '    If dtHeader.Rows.Count <> 0 Then

        '        'Dim dtRow As DataRow = dtHeader.Rows(0)
        '        With m_oOrder.Ordhead
        '            '.Ord_Type = Trim(dtRow.Item("Ord_Type").ToString)
        '            '.Cus_No = Trim(dtRow.Item("Cus_No").ToString)
        '            'Call .GetCustomerDefaultValues()
        '            '.Cus_Alt_Adr_Cd = Trim(dtRow.Item("Cus_Alt_Adr_Cd").ToString)
        '            '.Slspsn_No = Trim(dtRow.Item("Slspsn_No").ToString)
        '            '.Slspsn_Pct_Comm = dtRow.Item("Slspsn_Pct_Comm")


        '            .User_Def_Fld_4 = Trim(dtRow.Item("User_Def_Fld_4").ToString)
        '            .Cus_Email_Address = Trim(dtRow.Item("Email_Address").ToString)
        '            '.Mfg_Loc = Trim(dtRow.Item("Mfg_Loc").ToString)
        '            '.Ship_Via_Cd = Trim(dtRow.Item("Ship_Via_Cd").ToString)
        '            '.Ar_Terms_Cd = Trim(dtRow.Item("Ar_Terms_Cd").ToString)
        '            '.Profit_Center = Trim(dtRow.Item("Profit_Center").ToString)
        '            '.Dept = Trim(dtRow.Item("Dept").ToString)
        '            '.User_Def_Fld_2 = Trim(dtRow.Item("User_Def_Fld_2").ToString)
        '            '.Discount_Pct = Trim(dtRow.Item("Discount_Pct").ToString)
        '            .User_Def_Fld_1 = Trim(dtRow.Item("User_Def_Fld_1").ToString)
        '            '.User_Def_Fld_3 = Trim(dtRow.Item("User_Def_Fld_3").ToString)
        '            .User_Def_Fld_5 = Trim(dtRow.Item("User_Def_Fld_5").ToString)
        '            .Ship_Instruction_1 = Trim(dtRow.Item("Ship_Instruction_1").ToString)
        '            .Ship_Instruction_2 = Trim(dtRow.Item("Ship_Instruction_2").ToString)
        '            '.Bill_To_Name = Trim(dtRow.Item("Bill_To_Name").ToString)
        '            '.Bill_To_Addr_1 = Trim(dtRow.Item("Bill_To_Addr_1").ToString)
        '            '.Bill_To_Addr_2 = Trim(dtRow.Item("Bill_To_Addr_2").ToString)
        '            '.Bill_To_Addr_3 = Trim(dtRow.Item("Bill_To_Addr_3").ToString)
        '            '.Bill_To_Addr_4 = Trim(dtRow.Item("Bill_To_Addr_4").ToString)
        '            '.Bill_To_City = Trim(dtRow.Item("Bill_To_City").ToString)
        '            '.Bill_To_State = Trim(dtRow.Item("Bill_To_State").ToString)
        '            '.Bill_To_Zip = Trim(dtRow.Item("Bill_To_Zip").ToString)
        '            '.Bill_To_Country = Trim(dtRow.Item("Bill_To_Country").ToString)
        '            .Ship_To_Name = Trim(dtRow.Item("Ship_To_Name").ToString)
        '            .Ship_To_Addr_1 = Trim(dtRow.Item("Ship_To_Addr_1").ToString)
        '            .Ship_To_Addr_2 = Trim(dtRow.Item("Ship_To_Addr_2").ToString)
        '            .Ship_To_Addr_3 = Trim(dtRow.Item("Ship_To_Addr_3").ToString)
        '            .Ship_To_Addr_4 = Trim(dtRow.Item("Ship_To_Addr_4").ToString)
        '            .Ship_To_City = Trim(dtRow.Item("Ship_To_City").ToString)
        '            .Ship_To_State = Trim(dtRow.Item("Ship_To_State").ToString)
        '            .Ship_To_Zip = Trim(dtRow.Item("Ship_To_Zip").ToString)
        '            .Ship_To_Country = Trim(dtRow.Item("Ship_To_Country").ToString)

        '            '.Curr_Cd = Trim(dtRow.Item("Curr_Cd").ToString)
        '            .Orig_Trx_Rt = dtRow.Item("Curr_Trx_Rt")
        '            '.Curr_Trx_Rt = dtRow.Item("Curr_Trx_Rt")
        '            '.Form_No = dtRow.Item("Form_No")
        '            '.Deter_Rate_By = dtRow.Item("Deter_Rate_By")
        '            '.Tax_Fg = Trim(dtRow.Item("Tax_Fg").ToString)
        '            .Contact_1 = Trim(dtRow.Item("Contact_1").ToString)
        '            .Phone_Number = Trim(dtRow.Item("Phone_Number").ToString)
        '            .Fax_Number = Trim(dtRow.Item("Fax_Number").ToString)
        '            .Email_Address = Trim(dtRow.Item("Email_Address").ToString)

        '            .Save()

        '        End With

        '    End If

        '    ' second part - create order lines

        '    If m_oOrder.FromHistory Then

        '        strSql = _
        '        "SELECT * " & _
        '        "FROM   OELINHST_SQL WITH (Nolock) " & _
        '        "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

        '    Else

        '        strSql = _
        '        "SELECT * " & _
        '        "FROM   OEORDLIN_SQL WITH (Nolock) " & _
        '        "WHERE  Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND Ord_Type = 'O' "

        '    End If

        '    dtLines = db.DataTable(strSql)

        '    If dtLines.Rows.Count <> 0 And dtHeader.Rows.Count <> 0 Then
        '        For Each dtRow In dtLines.Rows

        '            Dim oOrdline As New cOrdLine
        '            oOrdline.Source = OEI_SourceEnum.Macola_ExactRepeat
        '            oOrdline.Line_Seq_No = dtRow.Item("Line_Seq_No")
        '            oOrdline.OEOrdLin_ID = dtRow.Item("Line_Seq_No")
        '            oOrdline.Item_No = dtRow.Item("Item_No").ToString
        '            oOrdline.Loc = dtRow.Item("Loc").ToString

        '            ' Qty to ship is set to qty ordered, because of overpicks.
        '            Try
        '                oOrdline.Qty_Ordered = dtRow.Item("Qty_Ordered")
        '            Catch oe_er As OEException
        '                'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
        '            End Try

        '            Try
        '                oOrdline.Qty_To_Ship = dtRow.Item("Qty_Ordered")
        '                oOrdline.Qty_To_Ship_Changed(oOrdline.Item_No, oOrdline.Loc, False)
        '                'If oOrdline.Qty_Prev_Bkord > 0 Then
        '                '    oOrdline.Calculate_ETA()
        '                'End If
        '            Catch oe_er As OEException
        '                'If oe_er.Cancel Or oe_er.ShowMessage Then Throw oe_er ' New OEException(oe_er.Message, oe_er)
        '            End Try

        '            ' Do not enter unit price, let calculate new price on date
        '            ' oOrdline.Unit_Price = dtRow.Item("Unit_Price").ToString
        '            oOrdline.Discount_Pct = dtRow.Item("Discount_Pct")
        '            oOrdline.Request_Dt = Now.Date
        '            oOrdline.Promise_Dt = Now.Date
        '            oOrdline.Req_Ship_Dt = Now.Date
        '            oOrdline.Item_Desc_1 = dtRow.Item("Item_Desc_1").ToString
        '            oOrdline.Item_Desc_2 = dtRow.Item("Item_Desc_2").ToString
        '            oOrdline.Extra_1 = dtRow.Item("Extra_1").ToString

        '            oOrdline.Source = OEI_SourceEnum.frmOrder
        '            oOrdline.Save()

        '        Next dtRow

        '    End If

        '    ' Third part - create traveler. We cannot create in lines
        '    ' because of kit items.
        '    strSql = "" & _
        '    "SELECT 	L.*, ISNULL(B.Seq_No, 0) AS Kit_No " & _
        '    "FROM 		OEI_OrdLin L " & _
        '    "LEFT JOIN	OEI_OrdBld B WITH (NoLock) ON L.Ord_Guid = B.Ord_Guid AND L.Item_Guid = B.Comp_Item_Guid " & _
        '    "WHERE      L.Ord_Guid = '" & Trim(m_oOrder.Ordhead.Ord_GUID) & "' " & _
        '    "ORDER BY   L.Line_No, B.Seq_No "

        '    dtLines = db.DataTable(strSql)

        '    If dtLines.Rows.Count <> 0 Then

        '        For Each dtRow In dtLines.Rows

        '            'If m_oOrder.OrdLines.Contains(dtRow.Item("Item_Guid").ToString) Then

        '            Dim dtItem As DataTable

        '            strSql = "" & _
        '            "SELECT 	ISNULL(RouteID, 0) AS RouteID " & _
        '            "FROM 		Exact_Traveler_Header WITH (Nolock) " & _
        '            "WHERE		Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND " & _
        '            "   		Line_No = " & dtRow.Item("OEOrdLin_ID") & " AND " & _
        '            "           Component_Seq_No = " & dtRow.Item("Kit_No") & " " & _
        '            "ORDER BY 	Ord_No, Line_No, Component_Seq_No "

        '            dtItem = db.DataTable(strSql)
        '            If dtItem.Rows.Count <> 0 Then

        '                strSql = _
        '                "UPDATE OEI_OrdLin " & _
        '                "SET    RouteID = " & dtItem.Rows(0).Item("RouteID") & " " & _
        '                "WHERE  Ord_Guid = '" & dtRow.Item("Ord_Guid") & "' AND " & _
        '                "       Item_Guid = '" & dtRow.Item("Item_Guid") & "'"

        '                db.Execute(strSql)
        '                'db.DBCommand.CommandText = strSql
        '                'db.DBCommand.ExecuteNonQuery()
        '                'db.DBCommand.CommandText = ""

        '                'Dim oImprint As New cImprint(Trim(m_oOrder.Ordhead.Ord_GUID), Trim(dtItem.Rows(0).Item("Item_Guid").ToString))
        '                ''oImprint.Ord_Type = dtRow.Item("Ord_Type").ToString
        '                ''oImprint.OE_Line_No = dtRow.Item("OE_Line_No").ToString
        '                ''oImprint.Comp_Seq_No = dtRow.Item("Comp_Seq_No").ToString
        '                'oImprint.Item_No = dtRow.Item("Item_No").ToString
        '                ''oImprint.Par_Item_No = dtRow.Item("Par_Item_No").ToString
        '                'oImprint.Imprint = dtRow.Item("Extra_1").ToString
        '                'oImprint.Location = dtRow.Item("Extra_2").ToString
        '                'oImprint.Color = dtRow.Item("Extra_3").ToString
        '                'If IsNumeric(dtRow.Item("Extra_4").ToString) Then
        '                '    oImprint.Num = dtRow.Item("Extra_4").ToString
        '                'Else
        '                '    oImprint.Num = 0
        '                'End If
        '                'oImprint.Packaging = dtRow.Item("Extra_5").ToString
        '                ''oImprint.Extra_6 = dtRow.Item("Extra_6").ToString
        '                ''oImprint.Extra_7 = dtRow.Item("Extra_7").ToString
        '                'oImprint.Refill = dtRow.Item("Extra_8").ToString
        '                'oImprint.Laser_Setup = dtRow.Item("Extra_9").ToString
        '                'oImprint.Industry = dtRow.Item("Industry").ToString
        '                'oImprint.Comments = dtRow.Item("Comment").ToString
        '                'oImprint.Special_Comments = "REPEAT ORDER # " & Trim(pstrOrd_No) ' dtRow.Item("Comment2").ToString
        '                'oImprint.Save()

        '            End If
        '            'End If

        '        Next dtRow

        '    End If


        '    ' fourt part - create order imprint data. We cannot create in lines
        '    ' because of kit items.

        '    'strSql = "" & _
        '    '"SELECT 	EF.*, ISNULL(L.Line_No, 0) AS OEOrdLin_ID, ISNULL(EF.Comp_Seq_No, 0) AS Kit_No " & _
        '    '"FROM 		OELinHst_Sql L WITH (Nolock) " & _
        '    '"LEFT JOIN	Exact_Traveler_Extra_Fields EF WITH (Nolock) ON L.Ord_No = EF.Ord_No AND L.Line_Seq_No = EF.OE_Line_No " & _
        '    '"LEFT JOIN	IMOrdHst_Sql B WITH (Nolock) ON EF.Ord_No = B.Ord_No AND EF.OE_Line_No = B.Line_No AND EF.Comp_Seq_No = B.Seq_No " & _
        '    '"WHERE		L.Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' " & _
        '    '"ORDER BY 	L.Ord_No, L.Line_Seq_No, EF.Comp_Seq_No "

        '    ' No need to redo, will use the same one as previous query

        '    'dtLines = db.DataTable(strSql)

        '    If dtLines.Rows.Count <> 0 Then

        '        For Each dtRow In dtLines.Rows

        '            Dim dtItem As DataTable

        '            'strSql = "" & _
        '            '"SELECT 	L.* " & _
        '            '"FROM 		OEI_OrdLin L WITH (NoLock) " & _
        '            '"LEFT JOIN	OEI_OrdBld B WITH (NoLock) ON L.Ord_Guid = B.Ord_Guid AND L.Item_Guid = B.Comp_Item_Guid " & _
        '            '"WHERE      L.Ord_Guid = '" & Trim(m_oOrder.Ordhead.Ord_GUID) & "' AND " & _
        '            '"           L.Line_Seq_No = " & dtRow.Item("OEOrdLin_ID") & " " & _
        '            'IIf(dtRow.Item("Kit_No") = 0, "", " AND B.Seq_No = " & dtRow.Item("Kit_No") & " ") & _
        '            '"ORDER BY   L.Line_No, B.Seq_No "

        '            strSql = "" & _
        '            "SELECT 	*, CASE WHEN ISNULL(Num_Impr_1, 0) = 0 THEN CONVERT(INT, ISNULL(Extra_4, '0')) ELSE ISNULL(Num_Impr_1, 0) END AS Comp_Num_Impr " & _
        '            "FROM 		Exact_Traveler_Extra_Fields WITH (Nolock) " & _
        '            "WHERE		Ord_No = '" & Trim(pstrOrd_No).PadLeft(8) & "' AND " & _
        '            "   		OE_Line_No = " & dtRow.Item("OEOrdLin_ID") & " AND " & _
        '            "           Comp_Seq_No = " & dtRow.Item("Kit_No") & " " & _
        '            "ORDER BY 	Ord_No, OE_Line_No, Comp_Seq_No "

        '            dtItem = db.DataTable(strSql)
        '            If dtItem.Rows.Count <> 0 Then

        '                Dim oImprint As New cImprint(Trim(m_oOrder.Ordhead.Ord_GUID), Trim(dtRow.Item("Item_Guid").ToString))
        '                'oImprint.Ord_Type = dtRow.Item("Ord_Type").ToString
        '                'oImprint.OE_Line_No = dtRow.Item("OE_Line_No").ToString
        '                'oImprint.Comp_Seq_No = dtRow.Item("Comp_Seq_No").ToString
        '                oImprint.SaveToDB = False
        '                oImprint.Item_No = dtItem.Rows(0).Item("Item_No").ToString
        '                'oImprint.Par_Item_No = dtRow.Item("Par_Item_No").ToString
        '                oImprint.Imprint = dtItem.Rows(0).Item("Extra_1").ToString
        '                oImprint.Location = dtItem.Rows(0).Item("Extra_2").ToString
        '                oImprint.Color = dtItem.Rows(0).Item("Extra_3").ToString
        '                If IsNumeric(dtItem.Rows(0).Item("Comp_Num_Impr").ToString) Then
        '                    oImprint.Num_Impr_1 = dtItem.Rows(0).Item("Comp_Num_Impr")
        '                Else
        '                    oImprint.Num_Impr_1 = 0
        '                End If
        '                If IsNumeric(dtItem.Rows(0).Item("Num_Impr_2").ToString) Then
        '                    oImprint.Num_Impr_2 = dtItem.Rows(0).Item("Num_Impr_2")
        '                Else
        '                    oImprint.Num_Impr_2 = 0
        '                End If
        '                If IsNumeric(dtItem.Rows(0).Item("Num_Impr_3").ToString) Then
        '                    oImprint.Num_Impr_3 = dtItem.Rows(0).Item("Num_Impr_3")
        '                Else
        '                    oImprint.Num_Impr_3 = 0
        '                End If
        '                oImprint.Packaging = dtItem.Rows(0).Item("Extra_5").ToString
        '                'oImprint.Extra_6 = dtRow.Item("Extra_6").ToString
        '                'oImprint.Extra_7 = dtRow.Item("Extra_7").ToString
        '                oImprint.Refill = dtItem.Rows(0).Item("Extra_8").ToString
        '                oImprint.Laser_Setup = dtItem.Rows(0).Item("Extra_9").ToString
        '                oImprint.Industry = dtItem.Rows(0).Item("Industry").ToString
        '                oImprint.Comments = dtItem.Rows(0).Item("Comment").ToString
        '                oImprint.Repeat_From_Ord_No = dtItem.Rows(0).Item("Ord_No").ToString
        '                oImprint.Repeat_From_ID = dtItem.Rows(0).Item("ID")
        '                oImprint.Special_Comments = "REPEAT ORDER # " & Trim(pstrOrd_No) ' dtRow.Item("Comment2").ToString

        '                oImprint.SaveToDB = True

        '                oImprint.Save()

        '            End If

        '        Next dtRow

        '    End If

        '    strSql = "DELETE FROM OEI_Order_Contacts WHERE ORD_GUID = '" & m_oOrder.Ordhead.Ord_GUID & "' "
        '    db.Execute(strSql)

        '    strSql = _
        '    "INSERT INTO OEI_Order_Contacts (Ord_Guid, DefContact, DefMethod, ContactType) " & _
        '    "SELECT 	'" & m_oOrder.Ordhead.Ord_GUID & "', DefContact, DefMethod, ContactType  " & _
        '    "FROM 		Exact_Traveler_Order_Customer_Contacts_New " & _
        '    "WHERE 		OrderNo = '" & Trim(pstrOrd_No) & "'"
        '    db.Execute(strSql)

        '    m_oOrder.FromHistory = False

        'Catch oe_er As OEException
        '    Throw oe_er ' New OEException(oe_er.Message, oe_er)
        'Catch er As Exception
        '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub

#End Region


    ' PROPERTIES xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

#Region "Public Order management procedures ###############################"

    '    Public Sub Fill(Optional ByVal pOrderSource As Order_Source = Order_Source.Macola) ' ByRef m_oOrder As cOrder)
    Public Sub Fill() ' ByRef m_oOrder As cOrder)

        blnLoading = True

        Try
            'm_oOrder = m_oOrder

            txtCus_No.Enabled = True
            cmdCustomerSearch.Enabled = txtCus_No.Enabled

            txtOrd_No.Enabled = True
            cmdOrderSearch.Enabled = txtOrd_No.Enabled
            cboOrd_Type.Enabled = txtOrd_No.Enabled
            cmdCopyOrderFromHistory.Enabled = txtOrd_No.Enabled

            'If Trim(m_oOrder.Ordhead.Ord_No).Length <> 0 Then
            If m_oOrder.Ordhead.Ord_Type Is Nothing Then
                cboOrd_Type.Text = "O=Order" ' m_oOrder.Ordhead.Ord_Type
            Else
                cboOrd_Type.Text = m_oOrder.Ordhead.Ord_Type
            End If

            If m_oOrder.OrderSource = OrderSourceEnum.Macola Then
                txtOrd_No.Text = m_oOrder.Ordhead.Ord_No
            ElseIf m_oOrder.OrderSource = OrderSourceEnum.OEInterface Then
                txtOrd_No.Text = m_oOrder.Ordhead.OEI_Ord_No
            End If

            'txtOrd_Dt.Text = String.Format("{0:0:MM/dd/yyyy}", m_oOrder.Ordhead.Ord_Dt)

            If Not (m_oOrder.Ordhead.Ord_Dt.Equals(NoDate())) Then
                txtOrd_Dt.Text = String.Format("{0:MM/dd/yyyy}", m_oOrder.Ordhead.Ord_Dt)
                If m_oOrder.Ordhead.Ord_Dt_Shipped.Equals(NoDate()) Then
                    txtOrd_Dt_Shipped.Text = "A.S.A.P."
                Else
                    'txtOrd_Dt_Shipped.Text = String.Format("{0:MM/dd/yyyy}", m_oOrder.Ordhead.Ord_Dt_Shipped)
                    txtOrd_Dt_Shipped.Text = String.Format("{0:MM/dd/yyyy}", m_oOrder.Ordhead.Shipping_Dt)
                End If
            Else
                txtOrd_Dt.Text = ""
                txtOrd_Dt_Shipped.Text = ""
            End If

            If Not (m_oOrder.Ordhead.In_Hands_Date.Equals(NoDate())) Then
                txtInHandsDate.Text = String.Format("{0:MM/dd/yyyy}", m_oOrder.Ordhead.In_Hands_Date)
            Else
                txtInHandsDate.Text = ""
            End If

            txtCus_No.Text = m_oOrder.Ordhead.Cus_No

            txtCustGroup.Text = m_oOrder.Ordhead.Customer.textfield3
            ttToolTip.SetToolTip(txtCustGroup, "Customer Group Name")

            If m_oOrder.Ordhead.Customer.TextField6 <> String.Empty Then
                lblVIP.Visible = True
            Else
                lblVIP.Visible = False
            End If

            txtCus_Alt_Adr_Cd.Text = IIf(Trim(m_oOrder.Ordhead.Cus_No) = "", "", IIf(Trim(m_oOrder.Ordhead.Cus_Alt_Adr_Cd) = Nothing, "SAME", m_oOrder.Ordhead.Cus_Alt_Adr_Cd))
            txtOe_Po_No.Text = m_oOrder.Ordhead.Oe_Po_No
            txtSlspsn_No.Text = IIf(m_oOrder.Ordhead.Cus_No = "", "", m_oOrder.Ordhead.Slspsn_No)


            '++ID 07142024 fill RE:SSP
            Call RPSSP()


            If m_oOrder.Ordhead.intSSP <> 0 Then
                cboreSSP.SelectedValue = m_oOrder.Ordhead.intSSP
            End If

            '++ID 07312024 inv_batch_id
            xTxtInvBatch.Text = m_oOrder.Ordhead.Inv_Batch_Id

            txtJob_No.Text = m_oOrder.Ordhead.Job_No



            txtMfg_Loc.Text = m_oOrder.Ordhead.Mfg_Loc




            txtShip_Via_Cd.Text = m_oOrder.Ordhead.Ship_Via_Cd
            txtAr_Terms_Cd.Text = m_oOrder.Ordhead.Ar_Terms_Cd
            txtDiscount_Pct.Text = 0 '++ID 7.5.2018 put in comment IIf(m_oOrder.Ordhead.Cus_No = "", "", m_oOrder.Ordhead.Discount_Pct)

            txtProfit_Center.Text = m_oOrder.Ordhead.Profit_Center
            txtDept.Text = m_oOrder.Ordhead.Dept

            txtShip_Instruction_1.Text = m_oOrder.Ordhead.Ship_Instruction_1
            txtShip_Instruction_2.Text = m_oOrder.Ordhead.Ship_Instruction_2

            txtBill_To_Name.Text = m_oOrder.Ordhead.Bill_To_Name
            txtBill_To_Addr_1.Text = m_oOrder.Ordhead.Bill_To_Addr_1
            txtBill_To_Addr_2.Text = m_oOrder.Ordhead.Bill_To_Addr_2
            txtBill_To_Addr_3.Text = m_oOrder.Ordhead.Bill_To_Addr_3
            txtBill_To_Addr_4.Text = m_oOrder.Ordhead.Bill_To_Addr_4
            txtBill_To_City.Text = m_oOrder.Ordhead.Bill_To_City
            txtBill_To_State.Text = m_oOrder.Ordhead.Bill_To_State
            txtBill_To_Zip.Text = m_oOrder.Ordhead.Bill_To_Zip
            txtBill_To_Country.Text = m_oOrder.Ordhead.Bill_To_Country

            txtShip_To_Name.Text = m_oOrder.Ordhead.Ship_To_Name
            txtShip_To_Addr_1.Text = m_oOrder.Ordhead.Ship_To_Addr_1
            txtShip_To_Addr_2.Text = m_oOrder.Ordhead.Ship_To_Addr_2
            txtShip_To_Addr_3.Text = m_oOrder.Ordhead.Ship_To_Addr_3
            txtShip_To_Addr_4.Text = m_oOrder.Ordhead.Ship_To_Addr_4
            txtShip_To_City.Text = m_oOrder.Ordhead.Ship_To_City
            txtShip_To_State.Text = m_oOrder.Ordhead.Ship_To_State
            txtShip_To_Zip.Text = m_oOrder.Ordhead.Ship_To_Zip
            txtShip_To_Country.Text = m_oOrder.Ordhead.Ship_To_Country
            'End If
            txtEnd_User.Text = m_oOrder.Ordhead.End_User
            txtUser_Def_Fld_1.Text = m_oOrder.Ordhead.User_Def_Fld_1
            cboUser_Def_Fld_2.SelectedText = IIf(m_oOrder.Ordhead.User_Def_Fld_2 Is Nothing, "", Trim(m_oOrder.Ordhead.User_Def_Fld_2))
            cboUser_Def_Fld_2.Text = IIf(m_oOrder.Ordhead.User_Def_Fld_2 Is Nothing, "", Trim(m_oOrder.Ordhead.User_Def_Fld_2))

            '++ID 06.17.2020 added new field only for oei_ordheader
            cbo_Extra_9.SelectedText = IIf(m_oOrder.Ordhead.Extra_9 Is Nothing, "", Trim(m_oOrder.Ordhead.Extra_9))
            cbo_Extra_9.Text = IIf(m_oOrder.Ordhead.Extra_9 Is Nothing, "", Trim(m_oOrder.Ordhead.Extra_9))
            '---------------------

            txtUser_Def_Fld_3.Text = m_oOrder.Ordhead.User_Def_Fld_3
            txtUser_Def_Fld_4.Text = m_oOrder.Ordhead.User_Def_Fld_4
            txtUser_Def_Fld_5.Text = m_oOrder.Ordhead.User_Def_Fld_5

            txtProg_Spector_Cd.Text = m_oOrder.Ordhead.Prog_Spector_Cd

            txtContactEmail.Text = m_oOrder.Ordhead.Email_Address

            'Call Set_Curr_Trx_Rt()

            'txtCus_No.Enabled = (Trim(txtCus_No.Text) = "")
            'cmdCustomerSearch.Enabled = txtCus_No.Enabled

            txtOrd_No.Enabled = (Trim(txtCus_No.Text) = "")
            cmdOrderSearch.Enabled = txtOrd_No.Enabled
            cboOrd_Type.Enabled = txtOrd_No.Enabled
            cmdCopyOrderFromHistory.Enabled = txtOrd_No.Enabled

            txtShip_Link.Text = m_oOrder.Ordhead.Ship_Link
            'chkOrderAckSaveOnly.Checked = (m_oOrder.Ordhead.OrderAckSaveOnly = 1)
            chkRecover.Checked = False ' (Trim(txtCus_No.Text) = "")
            chkRecover.Enabled = (Trim(txtCus_No.Text) = "")

            chkExactRepeat.Checked = False ' (Trim(txtCus_No.Text) = "")
            chkExactRepeat.Enabled = (Trim(txtCus_No.Text) = "")

            blnLoading = False

            If m_oOrder.Is_White_Flag() Then
                RaiseEvent White_Flag_Customer(True, New System.EventArgs())
            End If

            If txtCus_No.Text = "" Then txtOrd_No.Focus()

            '++ID ---------------------------- 05.02.2024----------------------
            'need to check if order already was entered and exist any saved record


            Dim oTr_Reas_notvegas As cTraveler_Reason_LateShipment
            oTr_Reas_notvegas = New cTraveler_Reason_LateShipment

            oTr_Reas_notvegas = oTr_Reas_notvegas.Load(m_oOrder.Ordhead.Ord_GUID)

            If Not oTr_Reas_notvegas Is Nothing Then
                If oTr_Reas_notvegas.REASON_TYPE_ID <> 0 Then
                    cboWhyNotVegas.SelectedValue = oTr_Reas_notvegas.REASON_TYPE_ID
                End If
            End If
            '-----------------------------------------------------------------



        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Reset() ' ByRef pOrder As cOrder) 
        Try
            cboOrd_Type.Text = "O=Order"
            txtOrd_No.Text = String.Empty
            txtOrd_Dt.Text = String.Empty

            txtApplyTo.Text = String.Empty
            txtOrd_Dt_Shipped.Text = String.Empty
            txtInHandsDate.Text = String.Empty
            txtCus_No.Text = String.Empty
            '++ID 8.15.2018
            txtCustGroup.Text = String.Empty
            '++ID 09122024 VIP label
            lblVIP.Visible = False
            txtOe_Po_No.Text = String.Empty
            txtSlspsn_No.Text = String.Empty
            txtShip_Link.Text = String.Empty
            txtJob_No.Text = String.Empty
            txtContactEmail.Text = String.Empty

            txtCus_Alt_Adr_Cd.Text = String.Empty
            txtMfg_Loc.Text = String.Empty
            txtShip_Via_Cd.Text = String.Empty
            txtAr_Terms_Cd.Text = String.Empty
            txtDiscount_Pct.Text = String.Empty
            txtProfit_Center.Text = String.Empty
            txtDept.Text = String.Empty
            txtShip_Instruction_1.Text = String.Empty
            txtShip_Instruction_2.Text = String.Empty

            txtBill_To_Name.Text = String.Empty
            txtBill_To_Addr_1.Text = String.Empty
            txtBill_To_Addr_2.Text = String.Empty
            txtBill_To_Addr_3.Text = String.Empty
            txtBill_To_Addr_4.Text = String.Empty
            txtBill_To_City.Text = String.Empty
            txtBill_To_State.Text = String.Empty
            txtBill_To_Zip.Text = String.Empty
            txtBill_To_Country.Text = String.Empty
            lblBillToCountryDesc.Text = String.Empty

            txtShip_To_Name.Text = String.Empty
            txtShip_To_Addr_1.Text = String.Empty
            txtShip_To_Addr_2.Text = String.Empty
            txtShip_To_Addr_3.Text = String.Empty
            txtShip_To_Addr_4.Text = String.Empty
            txtShip_To_City.Text = String.Empty
            txtShip_To_State.Text = String.Empty
            txtShip_To_Zip.Text = String.Empty
            txtShip_To_Country.Text = String.Empty
            lblShipToCountryDesc.Text = String.Empty

            txtEnd_User.Text = String.Empty
            txtUser_Def_Fld_1.Text = String.Empty
            cboUser_Def_Fld_2.SelectedText = String.Empty
            cboUser_Def_Fld_2.Text = String.Empty

            '++ID 06.17.2020 added new field only for oei_ordheader
            cbo_Extra_9.SelectedText = String.Empty
            cbo_Extra_9.Text = String.Empty
            '---------------------------

            txtUser_Def_Fld_3.Text = String.Empty
            txtUser_Def_Fld_4.Text = String.Empty
            txtUser_Def_Fld_5.Text = String.Empty
            txtProg_Spector_Cd.Text = String.Empty
            'txtOrd_Dt_Shipped.Text = String.Empty
            txtCus_No.Enabled = True
            cmdCustomerSearch.Enabled = txtCus_No.Enabled

            txtOrd_No.Enabled = True
            cmdOrderSearch.Enabled = txtOrd_No.Enabled
            cboOrd_Type.Enabled = txtOrd_No.Enabled
            cmdCopyOrderFromHistory.Enabled = txtOrd_No.Enabled
            'chkOrderAckSaveOnly.Checked = False
            chkRecover.Enabled = True
            chkRecover.Checked = False

            chkExactRepeat.Enabled = True
            chkExactRepeat.Checked = False

            Call SetEDIButton()

            cboreSSP.SelectedValue = 0
            '++ID 07312024
            xTxtInvBatch.Text = String.Empty

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Sub Save()

        Try
            If cboOrd_Type.SelectedItem Is Nothing Then
                m_oOrder.Ordhead.Ord_Type = "O=Order"
            Else
                m_oOrder.Ordhead.Ord_Type = cboOrd_Type.SelectedItem.ToString ' cboOrderType
            End If
            m_oOrder.Ordhead.Ord_No = txtOrd_No.Text '  txtOrderNo  
            m_oOrder.Ordhead.OEI_Ord_No = txtOrd_No.Text '  txtOrderNo  

            'm_oOrder.Ordhead.Ord_Dt = txtOrd_Dt.Text ' txtOrderDate
            If IsDate(txtOrd_Dt.Text) Then
                m_oOrder.Ordhead.Ord_Dt = txtOrd_Dt.Text ' txtShipDate
                If IsDate(txtOrd_Dt_Shipped.Text) Then
                    m_oOrder.Ordhead.Ord_Dt_Shipped = txtOrd_Dt_Shipped.Text ' txtShipDate
                    m_oOrder.Ordhead.Shipping_Dt = txtOrd_Dt_Shipped.Text ' txtShipDate
                Else
                    m_oOrder.Ordhead.Ord_Dt_Shipped = NoDate()
                    m_oOrder.Ordhead.Shipping_Dt = NoDate()
                End If
            Else
                m_oOrder.Ordhead.Ord_Dt = NoDate()
                m_oOrder.Ordhead.Ord_Dt_Shipped = NoDate()
            End If

            'Special case here, if date indicated then use it, else show 'A.S.A.P.' with NULL DATE
            If IsDate(txtOrd_Dt_Shipped.Text) Then
                m_oOrder.Ordhead.Ord_Dt_Shipped = txtOrd_Dt_Shipped.Text ' txtShipDate Shipping_Dt
                m_oOrder.Ordhead.Shipping_Dt = txtOrd_Dt_Shipped.Text ' txtShipDate Shipping_Dt
            Else
                m_oOrder.Ordhead.Ord_Dt_Shipped = NoDate()
                m_oOrder.Ordhead.Shipping_Dt = NoDate()
            End If

            m_oOrder.Ordhead.Cus_No = txtCus_No.Text ' txtCus_No
            m_oOrder.Ordhead.Cus_Alt_Adr_Cd = txtCus_Alt_Adr_Cd.Text ' txtCus_Alt_Adr_Cd
            m_oOrder.Ordhead.Oe_Po_No = txtOe_Po_No.Text ' txtOe_Po_No
            m_oOrder.Ordhead.Slspsn_No = IIf(IsNumeric(txtSlspsn_No.Text), txtSlspsn_No.Text, 0) ' txtSlspsn_No
            m_oOrder.Ordhead.Ship_Link = txtShip_Link.Text

            m_oOrder.Ordhead.Job_No = txtJob_No.Text ' txtJob_No 

            m_oOrder.Ordhead.Mfg_Loc = txtMfg_Loc.Text ' txtMfg_Loc
            m_oOrder.Ordhead.Ship_Via_Cd = txtShip_Via_Cd.Text ' txtShip_Via_Cd
            m_oOrder.Ordhead.Ar_Terms_Cd = txtAr_Terms_Cd.Text ' txtAr_Terms_Cd
            m_oOrder.Ordhead.Discount_Pct = 0 '++ID put in comment   IIf(IsNumeric(txtDiscount_Pct.Text), txtDiscount_Pct.Text, 0) ' txtDiscount_Pct

            m_oOrder.Ordhead.Profit_Center = txtProfit_Center.Text ' txtProfit_Center
            m_oOrder.Ordhead.Dept = txtDept.Text ' txtDept

            m_oOrder.Ordhead.Ship_Instruction_1 = txtShip_Instruction_1.Text ' txtShip_Instruction_1
            m_oOrder.Ordhead.Ship_Instruction_2 = txtShip_Instruction_2.Text ' txtShip_Instruction_2

            m_oOrder.Ordhead.Bill_To_Name = txtBill_To_Name.Text ' txtBill_To_Name
            m_oOrder.Ordhead.Bill_To_Addr_1 = txtBill_To_Addr_1.Text ' txtBill_To_Addr_1
            m_oOrder.Ordhead.Bill_To_Addr_2 = txtBill_To_Addr_2.Text ' txtBill_To_Addr_2
            m_oOrder.Ordhead.Bill_To_Addr_3 = txtBill_To_Addr_3.Text ' txtBill_To_Addr_3
            m_oOrder.Ordhead.Bill_To_Addr_4 = txtBill_To_Addr_4.Text ' txtBill_To_Addr_4
            m_oOrder.Ordhead.Bill_To_City = txtBill_To_City.Text ' txtBill_To_Addr_4
            m_oOrder.Ordhead.Bill_To_State = txtBill_To_State.Text ' txtBill_To_Addr_4
            m_oOrder.Ordhead.Bill_To_Zip = txtBill_To_Zip.Text ' txtBill_To_Addr_4
            m_oOrder.Ordhead.Bill_To_Country = txtBill_To_Country.Text ' txtBill_To_Country

            m_oOrder.Ordhead.Ship_To_Name = txtShip_To_Name.Text ' txtShip_To_Name
            m_oOrder.Ordhead.Ship_To_Addr_1 = txtShip_To_Addr_1.Text ' txtShip_To_Addr_1
            m_oOrder.Ordhead.Ship_To_Addr_2 = txtShip_To_Addr_2.Text ' txtShip_To_Addr_2
            m_oOrder.Ordhead.Ship_To_Addr_3 = txtShip_To_Addr_3.Text ' txtShip_To_Addr_3
            m_oOrder.Ordhead.Ship_To_Addr_4 = txtShip_To_Addr_4.Text ' txtShip_To_Addr_4
            m_oOrder.Ordhead.Ship_To_City = txtShip_To_City.Text ' txtBill_To_Addr_4
            m_oOrder.Ordhead.Ship_To_State = txtShip_To_State.Text ' txtBill_To_Addr_4
            m_oOrder.Ordhead.Ship_To_Zip = txtShip_To_Zip.Text ' txtBill_To_Addr_4
            m_oOrder.Ordhead.Ship_To_Country = txtShip_To_Country.Text ' txtShip_To_Country

            m_oOrder.Ordhead.End_User = txtEnd_User.Text
            m_oOrder.Ordhead.User_Def_Fld_1 = txtUser_Def_Fld_1.Text
            m_oOrder.Ordhead.User_Def_Fld_2 = cboUser_Def_Fld_2.Text

            '++ID 06.17.2020 added new field only for oei_ordheader
            m_oOrder.Ordhead.Extra_9 = cbo_Extra_9.Text
            '--------------------------------
            m_oOrder.Ordhead.User_Def_Fld_3 = txtUser_Def_Fld_3.Text
            m_oOrder.Ordhead.User_Def_Fld_4 = txtUser_Def_Fld_4.Text
            m_oOrder.Ordhead.User_Def_Fld_5 = txtUser_Def_Fld_5.Text
            m_oOrder.Ordhead.Prog_Spector_Cd = txtProg_Spector_Cd.Text

            m_oOrder.Ordhead.Email_Address = txtContactEmail.Text

            If IsDate(txtInHandsDate.Text) Then
                m_oOrder.Ordhead.In_Hands_Date = txtInHandsDate.Text ' txtShipDate Shipping_Dt
            Else
                m_oOrder.Ordhead.In_Hands_Date = NoDate()
            End If
            'm_oOrder.Ordhead.In_Hands_Date = txtInHandsDate.Text

            '07.04.2024  new SSP field ticket#40254
            m_oOrder.Ordhead.intSSP = cboreSSP.SelectedValue

            '++ID 07312024
            m_oOrder.Ordhead.Inv_Batch_Id = xTxtInvBatch.Text

            'm_oOrder.Ordhead.OrderAckSaveOnly = IIf(chkOrderAckSaveOnly.Checked, 1, 0)

            'm_oOrder.Ordhead.InvalidShipDateEmail = (chkInvalidShipDateEmail.Checked)
            'm_oOrder.Ordhead.InvalidEmail = (chkInvalidEmail.Checked)
            m_oOrder.Ordhead.Save()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

#End Region


#Region "Public properties ################################################"

    Public ReadOnly Property Ord_No() As String
        Get
            Ord_No = txtOrd_No.Text
        End Get
    End Property

    Public ReadOnly Property Order() As cOrder
        Get
            Order = m_oOrder
        End Get
    End Property

#End Region

    Private Sub cmdOEFlash_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOEFlash.Click

        Try

            'If m_oOrder.Ordhead.Cus_No Is Nothing Then Exit Sub
            'If Trim(m_oOrder.Ordhead.Cus_No) = "" Then Exit Sub

            'Dim ofrmOEFlash As New frmOEFlash(m_oOrder.Ordhead.Cus_No)

            'ofrmOEFlash.ShowDialog()

            'ofrmOEFlash = Nothing

            If m_oOrder.Ordhead.Cus_No Is Nothing Then Exit Sub
            If Trim(m_oOrder.Ordhead.Cus_No) = "" Then Exit Sub

            If ofrmOEFlash Is Nothing Then
                'Debug.Print("01")
                ofrmOEFlash = New frmOEFlash(m_oOrder.Ordhead.Cus_No)
                ofrmOEFlash.Show()
                'ofrmOEFlash = Nothing
            ElseIf Not (ofrmOEFlash.Visible) Then
                'Debug.Print("02")
                ofrmOEFlash = New frmOEFlash(m_oOrder.Ordhead.Cus_No)
                ofrmOEFlash.Visible = True
            Else
                'Debug.Print("03")
                ofrmOEFlash.Cus_No = m_oOrder.Ordhead.Cus_No
                'If ofrmOEFlash.Visible = False Then ofrmOEFlash.Show()
                ofrmOEFlash.Show()
                If ofrmOEFlash.dgvComments.Rows.Count <> 0 Then ofrmOEFlash.Focus()
            End If

            'ofrmOEFlash = Nothing

        Catch er As Exception
            MsgBox(er.Message)
        End Try

    End Sub

    Private Sub cmdShip_LinkEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShip_LinkEdit.Click

        Dim strShipLink As String
        If Trim(txtShip_Link.Text) = "" Then
            strShipLink = txtOrd_No.Text
        Else
            strShipLink = txtShip_Link.Text
        End If

        Dim oFormShipLinks As New frmShipLinks(txtCus_No.Text, txtOrd_No.Text, strShipLink)
        oFormShipLinks.ShowDialog()

        oFormShipLinks.Dispose()

    End Sub

    Private Sub txtOrd_No_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtOrd_No.Validating


        ''''''''''' PART TO USE WITH ONLY OEI LOAD
        'Try

        '    ' Checks for valid date format
        '    Call m_oOrder.Validation.OrderExists(txtOrd_No.Text)

        '    'txtCus_No.Tag = txtCus_No.Text
        '    'txtCus_No.OldValue = txtCus_No.Text

        'Catch Oe_er As OEException
        '    If Oe_er.Message = m_oOEMessage.Message(OEError.Order_No_Is_New) Then
        '        ' Do nothing
        '    ElseIf Oe_er.Message.ToString.Contains(m_oOEMessage.Message(OEError.Order_No_In_Use)) Then
        '        If chkRecover.Checked Then
        '            LoadOrderFromOEI(txtOrd_No.Text) ' Load order from OEI
        '            txtCus_No.Enabled = True ' (Trim(txtCus_No.Text) = "")
        '            cmdCustomerSearch.Enabled = txtCus_No.Enabled
        '            chkRecover.Checked = False
        '        Else
        '            e.Cancel = Oe_er.Cancel
        '            If Oe_er.ShowMessage Then MsgBox(Oe_er.Message)

        '            If e.Cancel Then
        '                txtOrd_No.Text = ""
        '                Exit Sub
        '            End If
        '        End If
        '    ElseIf Not (Oe_er.Message = m_oOEMessage.Message(OEError.Order_No_Not_Found)) Then
        '        e.Cancel = Oe_er.Cancel
        '        If Oe_er.ShowMessage Then MsgBox(Oe_er.Message)

        '        If e.Cancel Then
        '            txtOrd_No.Text = ""
        '            Exit Sub
        '        End If

        '        If Oe_er.Message <> m_oOEMessage.Message(OEError.Order_No_Found_In_OEI) Then Exit Sub

        '        LoadOrderFromOEI(txtOrd_No.Text) ' Load order from OEI
        '    Else
        '        e.Cancel = True
        '        MsgBox(Oe_er.Message)
        '        txtOrd_No.Text = ""
        '    End If
        'Catch er As Exception
        '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

        '''''''''' PART TO USE WITH OEI LOAD, EXACT REPEAT & REDO
        Try

            ' Checks for valid date format
            Call m_oOrder.Validation.OrderExists(txtOrd_No.Text)

            'txtCus_No.Tag = txtCus_No.Text
            'txtCus_No.OldValue = txtCus_No.Text

        Catch Oe_er As OEException
            If Oe_er.Message = m_oOEMessage.Message(OEError.Order_No_Is_New) Then
                ' Do nothing
            ElseIf Oe_er.Message.ToString.Contains(m_oOEMessage.Message(OEError.Order_No_In_Use)) Then

                If chkRecover.Checked Then

                    LoadOrderFromOEI(txtOrd_No.Text) ' Load order from OEI
                    txtCus_No.Enabled = True ' (Trim(txtCus_No.Text) = "")
                    cmdCustomerSearch.Enabled = txtCus_No.Enabled
                    chkRecover.Checked = False
                    chkExactRepeat.Checked = False

                Else

                    e.Cancel = Oe_er.Cancel
                    If Oe_er.ShowMessage Then MsgBox(Oe_er.Message)

                    If e.Cancel Then
                        txtOrd_No.Text = ""
                        Exit Sub
                    End If

                End If

            ElseIf (Oe_er.Message = m_oOEMessage.Message(OEError.Order_No_Found_In_History_Files_Entry_Not_Allowed)) Then

                If chkExactRepeat.Checked And chkRedo.Checked Then

                    e.Cancel = True
                    MsgBox("You must check only one of redo or exact repeat, not both.")
                    txtOrd_No.Text = ""
                    Exit Sub

                ElseIf chkExactRepeat.Checked Then

                    Try
                        CreateExactRepeatOrder(txtOrd_No.Text)
                        chkExactRepeat.Checked = False
                        txtOrd_No.Text = m_oOrder.Ordhead.OEI_Ord_No
                        'txtCus_Alt_Adr_Cd.Focus()
                        Dim oInvoke As New LoadOrderFromOEIDelegate(AddressOf LoadOrderFromOEI)
                        txtOrd_No.BeginInvoke(oInvoke, m_oOrder.Ordhead.OEI_Ord_No)

                    Catch Oe_er2 As OEException
                        If Oe_er2.ShowMessage Then MsgBox(Oe_er2.Message)
                        m_oOrder = New cOrder()
                        txtOrd_No.Text = ""
                        e.Cancel = True
                    End Try

                    'LoadOrderFromOEI(txtOrd_No.Text)

                ElseIf chkRedo.Checked Then

                    'CreateRedoOrder()
                    'txtOrd_No.Text = ""
                    'LoadOrderFromOEI()

                Else
                    e.Cancel = Oe_er.Cancel
                    If Oe_er.ShowMessage Then MsgBox(Oe_er.Message)

                    If e.Cancel Then
                        txtOrd_No.Text = ""
                        Exit Sub
                    End If
                End If

            ElseIf (Oe_er.Message = m_oOEMessage.Message(OEError.Order_No_From_Macola_Entry_Not_Allowed)) Then

                ' If order from exact repeat, we let go.
                'If chkExactRepeat.CheckState = CheckState.Checked Then Exit Sub ' m_oOrder.FromHistory
                'If m_oOrder.FromHistory Then Exit Sub

                If chkExactRepeat.Checked Then

                    Try
                        CreateExactRepeatOrder(txtOrd_No.Text)
                        chkExactRepeat.Checked = False
                        txtOrd_No.Text = m_oOrder.Ordhead.OEI_Ord_No
                        'txtCus_Alt_Adr_Cd.Focus()
                        Dim oInvoke As New LoadOrderFromOEIDelegate(AddressOf LoadOrderFromOEI)
                        txtOrd_No.BeginInvoke(oInvoke, m_oOrder.Ordhead.OEI_Ord_No)

                    Catch Oe_er2 As OEException
                        If Oe_er2.ShowMessage Then MsgBox(Oe_er2.Message)
                        m_oOrder = New cOrder()
                        txtOrd_No.Text = ""
                        e.Cancel = True
                    End Try

                    'LoadOrderFromOEI(txtOrd_No.Text)

                Else

                    e.Cancel = Oe_er.Cancel
                    If Oe_er.ShowMessage Then MsgBox(Oe_er.Message)

                    If e.Cancel Then
                        txtOrd_No.Text = ""
                        Exit Sub
                    End If

                    If Oe_er.Message <> m_oOEMessage.Message(OEError.Order_No_Found_In_OEI) Then Exit Sub

                    LoadOrderFromOEI(txtOrd_No.Text) ' Load order from OEI

                End If

            ElseIf (Oe_er.Message = m_oOEMessage.Message(OEError.OrderTypeInvoiceIsActive)) Then
                e.Cancel = Oe_er.Cancel
                If Oe_er.ShowMessage Then MsgBox(Oe_er.Message)

                If e.Cancel Then
                    txtOrd_No.Text = ""
                    Exit Sub
                End If

            ElseIf Not (Oe_er.Message = m_oOEMessage.Message(OEError.Order_No_Not_Found)) Then

                e.Cancel = Oe_er.Cancel
                If Oe_er.ShowMessage Then MsgBox(Oe_er.Message)

                If e.Cancel Then
                    txtOrd_No.Text = ""
                    Exit Sub
                End If

                If Oe_er.Message <> m_oOEMessage.Message(OEError.Order_No_Found_In_OEI) Then Exit Sub

                LoadOrderFromOEI(txtOrd_No.Text) ' Load order from OEI

            Else
                e.Cancel = True
                MsgBox(Oe_er.Message)
                txtOrd_No.Text = ""
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

        ''''''''DO NOT USE THIS PART
        'Catch Oe_er As OEException

        '' ''Try

        '' ''    'If txtCus_No.Text <> "" Then
        '' ''    If m_oOrder.Ordhead.Cus_No <> "" Then
        '' ''        Call CheckPendingOrder(OrderNoChangeType.ExistingOrder)
        '' ''    End If

        '' ''    'If Oe_er.Message = m_oOEMessage.Message(OEError.Order_No_Found_In_OEI) Then
        '' ''    '    LoadOrderFromOEI() ' Load order from OEI
        '' ''    'Else
        '' ''    '    MsgBox(Oe_er.Message)
        '' ''    'End If
        '' ''Catch Oe_er As OEException
        '' ''    e.Cancel = Oe_er.Cancel
        '' ''    If Oe_er.ShowMessage Then MsgBox(Oe_er.Message)
        '' ''    If Oe_er.Message = m_oOEMessage.Message(OEError.Order_No_Found_In_OEI) Then
        '' ''        Call Save()

        '' ''        'LoadOrderFromOEI() ' Load order from OEI
        '' ''    End If
        '' ''Catch ex As Exception
        '' ''    Throw New Exception(ex.Message, ex)
        '' ''End Try

        'Catch Oe_er As OEException
        '    e.Cancel = Oe_er.Cancel
        '    If Oe_er.ShowMessage Then MsgBox(Oe_er.Message)
        'Catch er As Exception
        '    MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        'End Try

    End Sub

    Private Sub txtCus_Alt_Adr_Cd_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCus_Alt_Adr_Cd.Validating

        'm_oOrder.Ordhead.GetAlternateAddress(txtCus_Alt_Adr_Cd.Text.ToString)

        'txtShip_To_Name.Text = m_oOrder.Ordhead.Ship_To_Name
        'txtShip_To_Addr_1.Text = m_oOrder.Ordhead.Ship_To_Addr_1
        'txtShip_To_Addr_2.Text = m_oOrder.Ordhead.Ship_To_Addr_2
        'txtShip_To_Addr_3.Text = m_oOrder.Ordhead.Ship_To_Addr_3
        'txtShip_To_Addr_4.Text = m_oOrder.Ordhead.Ship_To_Addr_4
        'txtShip_To_Country.Text = m_oOrder.Ordhead.Ship_To_Country

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Lines)

    End Sub

    'Private Sub ucOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

    '    ' If Ctl-Number, sets the new tab.
    '    If e.Control Then
    '        Select Case e.KeyCode
    '            Case Keys.D1
    '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Header)
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
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Salesperson)
    '            Case Keys.D8
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
    '            Case Keys.D9
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
    '        End Select

    '    End If

    'End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Try
            'Dim oForm As New frmItemPriceBreaks(m_oOrder.Ordhead.Curr_Cd, m_oOrder.Ordhead.Cus_No, g_oOrdline.Item_No, "", "")
            'Dim oForm As New frmItemLocQty("00EC1040RED", "1  ")
            Dim oForm As New frmItemInfo("00EC1040RED")
            oForm.ShowDialog()

            oForm.Dispose()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdComments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdComments.Click

        Try

            If m_oOrder.Ordhead.OEI_Ord_No Is Nothing Then Exit Sub
            If Trim(m_oOrder.Ordhead.OEI_Ord_No) = "" Then Exit Sub

            Dim ofrm As New cComments(m_oOrder.Ordhead.Ord_GUID)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdUser_Def_Fld_4Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUser_Def_Fld_4Search.Click

        If blnForceFocus Then Exit Sub

        Try

            Dim oSearch As New cSearch()
            oSearch.SearchElement = "OEI-REPEAT-CNT"
            oSearch.SetDefaultSearchElement("OEI_Repeat_Order_Contact_View.Cus_No", RTrim(txtCus_No.Text), "200")
            oSearch.Form.ShowDialog()
            If oSearch.Form.FoundElementValue <> String.Empty Then
                'Dim oSearch As New cSearch("AR-ALT-ADDR-COD")
                txtUser_Def_Fld_4.Text = oSearch.Form.FoundElementValue
                'cboShipTo.Text = txtCus_Alt_Adr_Cd.Text
            End If
            oSearch.Dispose()
            oSearch = Nothing
            txtUser_Def_Fld_4.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdProg_Spector_CdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdProg_Spector_CdSearch.Click


        If blnForceFocus Then Exit Sub

        Try

            Dim oSearch As New cSearch()
            oSearch.SearchElement = "OEI-PROGRAM"
            oSearch.SetDefaultSearchElement("MDB_CUS_PROG.Cus_No", RTrim(txtCus_No.Text), "200")
            oSearch.Form.ShowDialog()
            oSearch.Dispose()
            oSearch = Nothing
            txtProg_Spector_Cd.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdCopyDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyDate.Click

        Dim iPos As Integer = 0

        '18 = Request_Dt
        '19 = Promise_Dt
        '20 = Req_Ship_Dt

        'Dim oRequested As Date = txtInHandsDate.Text
        'Dim oPromise As Date = txtOrd_Dt_Shipped.Text

        Try
            If m_oOrder.OrdLines.Count = 0 Then Exit Sub

            iPos = 1
            If oFrmOrder.dgvItems.Rows.Count <> 0 Then
                iPos = 2
                For Each oOrdline As cOrdLine In m_oOrder.OrdLines
                    iPos = 3
                    If IsDate(txtInHandsDate.Text) Then
                        iPos = 4
                        oOrdline.Request_Dt = CDate(txtInHandsDate.Text)
                        iPos = 5
                        oOrdline.Save()

                    End If

                    If IsDate(txtOrd_Dt_Shipped.Text) Then
                        iPos = 4
                        oOrdline.Promise_Dt = CDate(txtOrd_Dt_Shipped.Text)
                        'oOrdline.Req_Ship_Dt = CDate(txtOrd_Dt_Shipped.Text)
                        iPos = 5
                        oOrdline.Save()

                    End If

                Next

                iPos = 11
                For Each dgvRow As DataGridViewRow In oFrmOrder.dgvItems.Rows
                    iPos = 12
                    If IsDate(txtOrd_Dt_Shipped.Text) Then
                        iPos = 13
                        'dgvRow.Cells(18).Value = CDate(txtOrd_Dt_Shipped.Text)
                        dgvRow.Cells(oFrmOrder.Columns.Request_Dt).Value = CDate(txtOrd_Dt_Shipped.Text)  'Changed July 28, 2017 - T. Louzon
                        iPos = 14
                        'dgvRow.Cells(19).Value = CDate(txtOrd_Dt_Shipped.Text)
                        dgvRow.Cells(oFrmOrder.Columns.Promise_Dt).Value = CDate(txtOrd_Dt_Shipped.Text)  'Changed July 28, 2017 - T. Louzon
                        iPos = 15
                        'dgvRow.Cells(20).Value = CDate(txtOrd_Dt_Shipped.Text)
                        dgvRow.Cells(oFrmOrder.Columns.Req_Ship_Dt).Value = CDate(txtOrd_Dt_Shipped.Text)  'Changed July 28, 2017 - T. Louzon

                    End If

                    If IsDate(txtInHandsDate.Text) Then
                        iPos = 13
                        'dgvRow.Cells(18).Value = CDate(txtInHandsDate.Text)
                        dgvRow.Cells(oFrmOrder.Columns.Request_Dt).Value = CDate(txtInHandsDate.Text)    'Changed July 28, 2017 - T. Louzon
                        'iPos = 14
                        'dgvRow.Cells(19).Value = CDate(txtOrd_Dt_Shipped.Text)
                        'iPos = 15
                        'dgvRow.Cells(20).Value = CDate(txtInHandsDate.Text)

                    End If

                Next dgvRow

                iPos = 21
                '==============================
                'Changed July 28, 2017 - T. Louzon
                'If oFrmOrder.dgvItems.CurrentRow.Cells(3).Equals(DBNull.Value) Then
                '    oFrmOrder.dgvItems.CurrentCell = oFrmOrder.dgvItems.Rows(0).Cells(17)
                'ElseIf oFrmOrder.dgvItems.CurrentRow.Cells(3).Value = "" Then
                '    oFrmOrder.dgvItems.CurrentCell = oFrmOrder.dgvItems.Rows(0).Cells(17)
                'End If
                If oFrmOrder.dgvItems.CurrentRow.Cells(oFrmOrder.Columns.Item_No).Equals(DBNull.Value) Then
                    oFrmOrder.dgvItems.CurrentCell = oFrmOrder.dgvItems.Rows(0).Cells(oFrmOrder.Columns.Discount_Pct)
                ElseIf oFrmOrder.dgvItems.CurrentRow.Cells(oFrmOrder.Columns.Item_No).Value = "" Then
                    oFrmOrder.dgvItems.CurrentCell = oFrmOrder.dgvItems.Rows(0).Cells(oFrmOrder.Columns.Discount_Pct)
                End If
                '==============================

                iPos = 22
                ' LAST REMOVED LINE
                g_oOrdline = m_oOrder.OrdLines(oFrmOrder.dgvItems.CurrentRow.Index + 1)

                'For Each oOrdline As cOrdLine In m_oOrder.OrdLines

                '    If IsDate(txtOrd_Dt_Shipped.Text) Then

                '        oOrdline.Promise_Dt = CDate(txtOrd_Dt_Shipped.Text)
                '        oOrdline.Save()

                '    End If

                'Next
                iPos = 23
                MsgBox("Date Changed.")

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " Pos : " & iPos.ToString)
        End Try

    End Sub

    'Private Sub cmdCopyDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyDate.Click

    '    Dim iPos As Integer = 0

    '    '18 = Request_Dt
    '    '19 = Promise_Dt
    '    '20 = Req_Ship_Dt

    '    Try
    '        If m_oOrder.OrdLines.Count = 0 Then Exit Sub

    '        iPos = 1
    '        If oFrmOrder.dgvItems.Rows.Count <> 0 Then
    '            iPos = 2
    '            For Each oOrdline As cOrdLine In m_oOrder.OrdLines
    '                iPos = 3
    '                If IsDate(txtOrd_Dt_Shipped.Text) Then
    '                    iPos = 4
    '                    oOrdline.Request_Dt = CDate(txtOrd_Dt_Shipped.Text)
    '                    iPos = 5
    '                    oOrdline.Save()

    '                End If

    '                If IsDate(txtInHandsDate.Text) Then
    '                    iPos = 4
    '                    oOrdline.Req_Ship_Dt = CDate(txtInHandsDate.Text)
    '                    iPos = 5
    '                    oOrdline.Save()

    '                End If

    '            Next

    '            iPos = 11
    '            For Each dgvRow As DataGridViewRow In oFrmOrder.dgvItems.Rows
    '                iPos = 12
    '                If IsDate(txtOrd_Dt_Shipped.Text) Then
    '                    iPos = 13
    '                    dgvRow.Cells(18).Value = CDate(txtOrd_Dt_Shipped.Text)
    '                    iPos = 14
    '                    dgvRow.Cells(19).Value = CDate(txtOrd_Dt_Shipped.Text)
    '                    iPos = 15
    '                    dgvRow.Cells(20).Value = CDate(txtOrd_Dt_Shipped.Text)

    '                End If

    '                If IsDate(txtInHandsDate.Text) Then
    '                    'iPos = 13
    '                    'dgvRow.Cells(18).Value = CDate(txtOrd_Dt_Shipped.Text)
    '                    'iPos = 14
    '                    'dgvRow.Cells(19).Value = CDate(txtOrd_Dt_Shipped.Text)
    '                    iPos = 15
    '                    dgvRow.Cells(20).Value = CDate(txtInHandsDate.Text)

    '                End If

    '            Next dgvRow

    '            iPos = 21
    '            If oFrmOrder.dgvItems.CurrentRow.Cells(3).Equals(DBNull.Value) Then
    '                oFrmOrder.dgvItems.CurrentCell = oFrmOrder.dgvItems.Rows(0).Cells(17)
    '            ElseIf oFrmOrder.dgvItems.CurrentRow.Cells(3).Value = "" Then
    '                oFrmOrder.dgvItems.CurrentCell = oFrmOrder.dgvItems.Rows(0).Cells(17)
    '            End If

    '            'oFrmOrder.dgvItems.CurrentCell = oFrmOrder.dgvItems.CurrentRow.Cells(17)
    '            iPos = 22
    '            ' LAST REMOVED LINE
    '            g_oOrdline = m_oOrder.OrdLines(oFrmOrder.dgvItems.CurrentRow.Index + 1)

    '            'For Each oOrdline As cOrdLine In m_oOrder.OrdLines

    '            '    If IsDate(txtOrd_Dt_Shipped.Text) Then

    '            '        oOrdline.Promise_Dt = CDate(txtOrd_Dt_Shipped.Text)
    '            '        oOrdline.Save()

    '            '    End If

    '            'Next
    '            iPos = 23
    '            MsgBox("Date Changed.")

    '        End If

    '    Catch er As Exception
    '        MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message & " Pos : " & iPos.ToString)
    '    End Try

    'End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Try
            'Dim dt As DataTable
            Dim db As New cDBA
            Dim strSql As String

            strSql = "EXEC DBO.OEI_IMPORT_TRIGGER '" & m_oOrder.Ordhead.Ord_GUID & "' "

            'dt = db.DataTable(strSql)
            db.Execute(strSql)

            'If dt.Rows.Count <> 0 Then
            '    MsgBox(dt.Rows(0).Item(0).ToString)
            'End If

            MsgBox("1")

        Catch er As Exception
            MsgBox(er.Message)
        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim dtLastCheckOrders As Date = Date.Now

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        'Call SendOrderConfirmationToUsers()

    End Sub



    'Public Sub SendOrderConfirmationToUsers(ByVal pstrOrd_Guid As String, ByVal pstrEmail As String, ByVal pTrigger As Trigger)
    'Public Sub SendOrderConfirmationToUsers()


    '    Dim dtUsers As DataTable
    '    Dim dt As DataTable
    '    Dim db As New cDBA

    '    Dim strSql As String = _
    '    "SELECT     DISTINCT ISNULL(O.UserID, '') AS UserID " & _
    '    "FROM       OEI_OrdHdr O WITH (Nolock) " & _
    '    "INNER JOIN	HumRes H WITH (Nolock) ON O.UserID = H.Usr_ID " & _
    '    "WHERE      O.TriggerTS IS NOT NULL AND O.Email_Sent = 0 "

    '    Try
    '        dtUsers = db.DataTable(strSql)

    '        If dtUsers.Rows.Count <> 0 Then

    '            For Each dtUserRow As DataRow In dtUsers.Rows

    '                strSql = _
    '                "SELECT     GETDATE() AS SENT_DATE, ISNULL(O.Cus_No, '') AS Cus_No, ISNULL(O.OE_Po_No, '') AS OE_Po_No, " & _
    '                "			ISNULL(O.OEI_Ord_No, '') AS OEI_Ord_No, ISNULL(O.Ord_No, '') AS Ord_No, " & _
    '                "			ISNULL(O.UserID, '') AS UserID, ISNULL(H.Mail, '') AS Email " & _
    '                "FROM       OEI_OrdHdr O WITH (Nolock) " & _
    '                "INNER JOIN	HumRes H WITH (Nolock) ON O.UserID = H.Usr_ID " & _
    '                "WHERE      TriggerTS IS NOT NULL AND O.Email_Sent = 0 AND O.UserID = '" & Trim(SqlCompliantString(dtUserRow.Item("UserID").ToString)) & "' "

    '                dt = db.DataTable(strSql)

    '                If dt.Rows.Count <> 0 Then

    '                    Call CreateEmailConnection()

    '                    mailer.Message.ToAddr = Trim(dt.Rows(0).Item("Email").ToString) ' "modification@spectorandco.ca"
    '                    mailer.Message.FromAddr = "Order Entry Interface <marcb@spectorandco.ca>"
    '                    mailer.Message.BCCAddr = "Email Log <email_log@bankers.ca>"
    '                    mailer.message.ccaddr = "marcbeauregard@yahoo.com"
    '                    mailer.Message.Subject = "Macola order importation report "
    '                    mailer.Message.BodyText = "<p style=""font-family:Courier"">" & "User ----------".PadRight(16) & "Customer No ---".PadRight(16) & "PO No ---------".PadRight(16) & "OEI Ord No ----".PadRight(16) & "Macola Ord No --".PadRight(16) & "<br>"

    '                    For Each dtRow As DataRow In dt.Rows
    '                        mailer.Message.BodyText &= _
    '                        (Trim(dtRow.Item("UserID").ToString)).PadRight(15, ".") & " " & _
    '                        (Trim(dtRow.Item("Cus_No").ToString)).PadRight(15, ".") & " " & _
    '                        (Trim(dtRow.Item("OE_Po_No").ToString)).PadRight(15, ".") & " " & _
    '                        (Trim(dtRow.Item("OEI_Ord_No").ToString)).PadRight(15, ".") & " " & _
    '                        (Trim(dtRow.Item("Ord_No").ToString)).PadRight(15) & "<br>"
    '                    Next dtRow

    '                    mailer.Message.BodyText &= "</p>"
    '                    mailer.Send()

    '                    Call CloseEmailConnection()

    '                    'strSql = _
    '                    '"UPDATE     OEI_OrdHdr SET Email_Sent = 1" & _
    '                    '"WHERE      TriggerTS IS NOT NULL AND Email_Sent = 0 AND " & _
    '                    '"           UserID = '" & Trim(SqlCompliantString(dtUserRow.Item("UserID").ToString)) & "' "

    '                    'db.Execute(strSql)

    '                End If

    '            Next dtUserRow

    '        End If

    '        'mailer.Message.Subject = strSubject
    '        'mailer.Message.BodyText = strMessage

    '    Catch er As Exception
    '    End Try

    'End Sub

    Public Sub CreateEmailConnection()
        Try
            'mailer = New MailBee.SMTP
            mailer = CreateObject("MailBee.SMTP")

            ' Unlock MailBee.SMTP object
            mailer.LicenseKey = "MBC500-5959843B6F-3D8F3ABAB139E365D283A4325B857897"

            ' SMTP server name
            mailer.ServerName = "uranium"

            ' Mark that body has HTML format
            mailer.Message.BodyFormat = 1

            ' Enable logging SMTP session into a file. If any errors  
            ' occur, the log can be used to investigate the problem.
            mailer.EnableLogging = True
            mailer.LogFilePath = "C:\HV_Control_SendLog.txt"

        Catch er As Exception

        End Try

    End Sub

    Public Sub CloseEmailConnection()

        Try
            ' Close the SMTP session
            If Not (mailer Is Nothing) Then mailer.Disconnect()

            'UPGRADE_NOTE: Object mailer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            mailer = Nothing


        Catch er As Exception

        End Try

    End Sub

    Private Sub cmdCustomerConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustomerConfig.Click

        Try

            If m_oOrder.Ordhead.Cus_No Is Nothing Then Exit Sub
            If Trim(m_oOrder.Ordhead.Cus_No) = "" Then Exit Sub

            If ofrmCustomerConfig Is Nothing Then
                ofrmCustomerConfig = New frmCustomerConfig(m_oOrder.Ordhead.Cus_No)
                ofrmCustomerConfig.Show()
                'ofrmOEFlash = Nothing
            ElseIf Not (ofrmCustomerConfig.Visible) Then
                ofrmCustomerConfig = New frmCustomerConfig(m_oOrder.Ordhead.Cus_No)
                ofrmCustomerConfig.Visible = True
            Else

                'Debug.Print("03")
                ofrmCustomerConfig.Cus_No = m_oOrder.Ordhead.Cus_No
                'If ofrmOEFlash.Visible = False Then ofrmOEFlash.Show()
                ofrmCustomerConfig.Show()
                If ofrmOEFlash.dgvComments.Rows.Count <> 0 Then ofrmOEFlash.Focus()
            End If

        Catch er As Exception
            MsgBox(er.Message)
        End Try

    End Sub

    Private Sub cmdOrderRequest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOrderRequest.Click


        Try
            Dim oSearch As New cSearch("OEI-INQUIRY")

            oSearch.Dispose()
            oSearch = Nothing

            'End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdSendEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSendEmail.Click


        Call SendOrderConfirmationForEDI("01EE3D19-0EE4-4CA1-A51F-05964E")

        'Call CreateEmailConnection()

        'mailer.Message.ToAddr = "marcbeauregard@yahoo.com"
        'mailer.Message.FromAddr = "marcb@spectorandco.ca"
        'mailer.Message.CCAddr = "sdfgsdfgsdfgsdfgsdfgjryukserdf@bankers.ca"

        'mailer.Message.Subject = "This is a test, don't reply to this email."
        'mailer.Message.BodyText = "This is a test, don't reply to this email."

        'If Mid(mailer.message.subject, 1, 5) = "*****" Then mailer.Message.BCCAddr = "marcb@spectorandco.ca"

        'mailer.Send()

        'Call CloseEmailConnection()

    End Sub


    Public Sub SendOrderConfirmationForEDI(ByVal pstrOrd_Guid As String)

        Dim dt As DataTable
        Dim db As New cDBA

        Dim strSql As String = _
        "SELECT		ISNULL(H.OEI_Ord_No, '') AS OEI_Ord_No, " & _
        "			ISNULL(H.Cus_No, '') AS Cus_No, " & _
        "			ISNULL(H.Bill_To_Name, '') AS Bill_To_Name, " & _
        "			ISNULL(H.Ord_No, '') AS Ord_No, " & _
        "			ISNULL(H.OE_PO_No, '') AS OE_PO_No, " & _
        "			CONVERT(SMALLDATETIME, ISNULL(H.Ord_Dt, ''), 103 ) AS Ord_Dt, " & _
        "			ISNULL(L.User_Def_Fld_4, '') AS User_Def_Fld_4, " & _
        "			ISNULL(L.Qty_Ordered, 0) AS Qty_Ordered, " & _
        "           ISNULL(U.Mail, 0) AS CSR_Mail " & _
        "FROM		OEI_OrdHdr H WITH (Nolock) " & _
        "INNER JOIN	OEI_ORDEDI E WITH (NOLOCK) ON H.Ord_Guid = E.Ord_Guid " & _
        "INNER JOIN	OEI_OrdLin L WITH (Nolock) ON E.Ord_Guid = L.Ord_Guid AND ISNULL(L.User_Def_Fld_4, '') <> '' " & _
        "           LEFT JOIN	HUMRES U WITH (Nolock) ON H.Slspsn_No = U.Res_ID " & _
        "WHERE		H.Ord_Guid = '" & Trim(pstrOrd_Guid) & "' "

        dt = db.DataTable(strSql)

        If dt.Rows.Count <> 0 Then
            Call CreateEmailConnection()

            mailer.Message.FromAddr = "EDI Service <it@spectorandco.ca>"
            mailer.Message.ToAddr = dt.Rows(0).Item("CSR_Mail").ToString.Trim
            mailer.Message.CCAddr = "clayton@spectorandco.com"

            mailer.Message.Subject = "New EDI Order for P.O. " & dt.Rows(0).Item("OE_PO_No").ToString.Trim & " from " & dt.Rows(0).Item("Bill_To_Name").ToString.Trim & ". "
            mailer.Message.BodyText = _
            "Customer number    : " & dt.Rows(0).Item("Cus_No").ToString.Trim & vbCrLf & _
            "Customer name      : " & dt.Rows(0).Item("Bill_To_Name").ToString.Trim & vbCrLf & _
            "Order number       : " & dt.Rows(0).Item("Ord_No").ToString.Trim & vbCrLf & _
            "P.O. number        : " & dt.Rows(0).Item("OE_PO_No").ToString.Trim & vbCrLf & _
            "Date               : " & dt.Rows(0).Item("Ord_Dt").ToString.Trim & vbCrLf & _
            "Decorated item     : " & dt.Rows(0).Item("User_Def_Fld_4").ToString.Trim & vbCrLf & _
            "Quantity ordered   : " & dt.Rows(0).Item("Qty_Ordered").ToString.Trim & vbCrLf

            mailer.Send()

            Call CloseEmailConnection()

        End If

    End Sub

    Private Sub SetEDIButton()

        Try

            Dim db As New cDBA()
            Dim dt As DataTable
            Dim strSql As String

            If g_User.EDI_Access = 1 Then

                strSql = "SELECT COUNT(ORDEDI_ID) AS ORD_QTY FROM OEI_OrdEDI WHERE OEI_Status = 'R' "
                dt = db.DataTable(strSql)

                'cmdEDI.Visible = (dt.Rows.Count <> 0)
                cmdEDI.Visible = (dt.Rows(0).Item("ORD_QTY") <> 0)

                If cmdEDI.Visible Then
                    qtyEdi = dt.Rows(0).Item("ORD_QTY")
                Else
                    qtyEdi = 0
                End If
            Else
                cmdEDI.Visible = False
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdEDI_MouseHover(sender As Object, e As System.EventArgs) Handles cmdEDI.MouseHover
        cmdEDI.Text = "(" & qtyEdi.ToString & ")"
    End Sub

    Private Sub cmdEDI_MouseLeave(sender As Object, e As System.EventArgs) Handles cmdEDI.MouseLeave
        cmdEDI.Text = "EDI"
    End Sub



    'Private Sub ucOrder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    '    Call SetEDIButton()

    'End Sub

    Private Sub cmdEDI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEDI.Click

        If txtOrd_No.Text <> "" Then
            MsgBox("You cannot call the EDI order in this state. You must finish with this order first.")
            Exit Sub
        End If

        Try

            Dim db As New cDBA()
            Dim dt As DataTable

            Dim strSql As String = _
            "SELECT TOP 1 * " & _
            "FROM   OEI_OrdEDI WITH (Nolock) " & _
            "WHERE  OEI_Status = 'R' "

            dt = db.DataTable(strSql)

            If dt.Rows.Count = 0 Then

                MsgBox("No EDI order to do.")
                Call SetEDIButton()
                Exit Sub

            End If

            m_oOrder = New cOrder()
            m_oOrder.LoadOrderFromXmlFile(dt.Rows(0).Item("OrdEDI_ID"))

            'CreateEDIOrderFromXML()
            txtOrd_No.Text = m_oOrder.Ordhead.OEI_Ord_No
            'txtCus_Alt_Adr_Cd.Focus()
            Dim oInvoke As New LoadOrderFromOEIDelegate(AddressOf LoadOrderFromOEI)
            txtOrd_No.BeginInvoke(oInvoke, m_oOrder.Ordhead.OEI_Ord_No)

            cmdEDI.Visible = False

        Catch Oe_er2 As OEException
            If Oe_er2.ShowMessage Then MsgBox(Oe_er2.Message)
            m_oOrder = New cOrder()
            txtOrd_No.Text = ""
        End Try

    End Sub

    Private Sub ucOrder_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus


        Call SetEDIButton()
        If txtCus_No.Text = "" Then txtOrd_No.Focus()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click


        'Dim db As New cDBA
        'Dim dt As DataTable
        'Dim strsql As String
        Dim oLogo As cMDB_Logo
        oLogo = New cMDB_Logo(1)

        oLogo.Logo_ID = 0
        oLogo.Logo_Name = "Copie de Rangers Jersey"
        oLogo.File_Url = "Rangers Copie.jpg"
        oLogo.Save()

    End Sub

    Private Sub cmdCustomerDash_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustomerDash.Click

        Try

            Dim startInfo As New ProcessStartInfo
            startInfo.FileName = "M:\mcc\traveler executable\CustomerFile\CustomerFile.application"
            'startInfo.FileName = "C:\dev\PUBLISH\Tools\Application Files\Tools_1_0_0_2\Tools.exe"
            'startInfo.Arguments = m_oOrder.Ordhead.Cus_No
            SaveSetting("Orders", "Startup", "Cus_no", m_oOrder.Ordhead.Cus_No.ToString)
            Process.Start(startInfo)

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub txtProg_Spector_Cd_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtProg_Spector_Cd.Validating

        Try

            Dim strOrd_No As String = m_oOrder.OneShotOrderUsed(txtProg_Spector_Cd.Text)
            If strOrd_No <> String.Empty Then
                MsgBox("This program is a one shot program and was already used on order " & strOrd_No.Trim & ".")
                txtProg_Spector_Cd.Text = String.Empty
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub ucOrder_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            IdUser = getUserID()
            xml_Access = getXMLAccess(IdUser)



            Dim db As New cDBA
            Dim dt As DataTable
            Dim strsql As String

            strsql = "select * from ( select '' As ENUM_VALUE " _
                    & " union all " _
                    & " select ENUM_VALUE  from mdb_cfg_enum where Enum_cat = 'OEI_EXTRA_9' " _
                    & " ) as t order by ENUM_VALUE"

            dt = db.DataTable(strsql)

            If dt.Rows.Count <> 0 Then
                cbo_Extra_9.DataSource = dt
                cbo_Extra_9.DisplayMember = "ENUM_VALUE"
                cbo_Extra_9.ValueMember = "ENUM_VALUE"
            End If

            '-----------------------------------load resaon why not Vegas-----------------------------------------------

            Dim enumTraveler_Reason_Type_LateShipment As cTraveler_Reason_Type_LateShipment

            Dim oTraveler_Reason_Type_LateShipment As cTraveler_Reason_Type_LateShipment
            oTraveler_Reason_Type_LateShipment = New cTraveler_Reason_Type_LateShipment

            Dim lstTraveler_Reason_Type_LateShipment As List(Of cTraveler_Reason_Type_LateShipment)
            lstTraveler_Reason_Type_LateShipment = New List(Of cTraveler_Reason_Type_LateShipment)

            lstTraveler_Reason_Type_LateShipment = oTraveler_Reason_Type_LateShipment.LoadLST

            enumTraveler_Reason_Type_LateShipment = New cTraveler_Reason_Type_LateShipment
            lstTraveler_Reason_Type_LateShipment.Insert(0, enumTraveler_Reason_Type_LateShipment)

            cboWhyNotVegas.DataSource = lstTraveler_Reason_Type_LateShipment
            cboWhyNotVegas.DisplayMember = "REASONTYPE"
            cboWhyNotVegas.ValueMember = "ID"

            'for existing line info will be filled in Function Fill from the same file ucOrders

            '-----------------------------------------------------------------------------------

            ''-----------------------------------load resaon RE:SSP----------------------------


            'Dim enumExact_Traveler_Rep_Codes As cExact_Traveler_Rep_Codes

            'Dim oExact_Traveler_Rep_Codes As cExact_Traveler_Rep_Codes
            'oExact_Traveler_Rep_Codes = New cExact_Traveler_Rep_Codes

            'Dim lstExact_Traveler_Rep_Codes As List(Of cExact_Traveler_Rep_Codes)
            'lstExact_Traveler_Rep_Codes = New List(Of cExact_Traveler_Rep_Codes)

            'lstExact_Traveler_Rep_Codes = oExact_Traveler_Rep_Codes.LoadLST

            'enumExact_Traveler_Rep_Codes = New cExact_Traveler_Rep_Codes
            'lstExact_Traveler_Rep_Codes.Insert(0, enumExact_Traveler_Rep_Codes)

            'cboWhyNotVegas.DataSource = lstExact_Traveler_Rep_Codes
            'cboWhyNotVegas.DisplayMember = "repname"
            'cboWhyNotVegas.ValueMember = "ID"


        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub RPSSP()
        Try
            '-----------------------------------load resaon RE:SSP----------------------------
            '     If _repcust = "" Then Exit Sub

            Dim enumExact_Traveler_Rep_Codes As cExact_Traveler_Rep_Codes

            Dim oExact_Traveler_Rep_Codes As cExact_Traveler_Rep_Codes
            oExact_Traveler_Rep_Codes = New cExact_Traveler_Rep_Codes

            Dim lstExact_Traveler_Rep_Codes As List(Of cExact_Traveler_Rep_Codes)
            lstExact_Traveler_Rep_Codes = New List(Of cExact_Traveler_Rep_Codes)

            lstExact_Traveler_Rep_Codes = oExact_Traveler_Rep_Codes.LoadLST()

            enumExact_Traveler_Rep_Codes = New cExact_Traveler_Rep_Codes
            lstExact_Traveler_Rep_Codes.Insert(0, enumExact_Traveler_Rep_Codes)

            cboreSSP.DataSource = lstExact_Traveler_Rep_Codes
            cboreSSP.DisplayMember = "repname"
            cboreSSP.ValueMember = "ID"

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Private Sub ucOrder_LostFocus(sender As Object, e As System.EventArgs) Handles Me.LostFocus
        Debug.Print("lOSTfOCUS")
    End Sub

    Private Sub cmdConvertQuote_Click(sender As System.Object, e As System.EventArgs) Handles cmdConvertQuote.Click

        Try
            Dim frmAB As New frmAddressBook()

            frmAB.ShowDialog()
            'End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub

    Function getUserID() As String
        'Get ID of the current user.
        Dim strSql As String
        Dim dt As New DataTable
        Dim db As New cDBA

        '  strSql = "SELECT ID FROM HUMRES WITH (NOLOCK) where fullname = 'stephane' "

        strSql = "SELECT ID FROM HUMRES WITH (NOLOCK) where fullname = '" + Environment.UserName + "' "
        dt = db.DataTable(strSql)
        If dt.Rows.Count <> 0 Then
            getUserID = dt.Rows(0).Item("ID").ToString.Trim
        Else
            getUserID = ""
        End If

        'strSql = "SELECT ID FROM HUMRES where fullname = '" + mGlobal.GetNTUserID + "' "
        'dt = db.DataTable(strSql)
        'Return Trim(dt.Rows(0).Item("ID"))
    End Function

    Function getXMLAccess(idUser As String) As String
        'Get XML_Access value of current user
        Dim strSql As String
        Dim dt As New DataTable
        Dim db As New cDBA

        Dim _return As String = ""

        strSql = "select XML_ACCESS from OEI_Users where HUMRES_ID = '" + idUser + "'"
        dt = db.DataTable(strSql)

        '++ID added if exeption for specify value (nothing vs empty or data)
        If dt.Rows.Count <> 0 Then
            _return = Trim(dt.Rows(0).Item("XML_Access")).ToString
        End If

        Return _return 'Trim(dt.Rows(0).Item("XML_Access"))

    End Function

    Private Sub cmdXml_Click(sender As System.Object, e As System.EventArgs) Handles cmdXml.Click
        Try

            Dim cord As New cOrder()
            cord.XML_Export(txtOrd_No.Text)
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub



    Private Sub cboUser_Def_Fld_2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboUser_Def_Fld_2.SelectedIndexChanged
        cOrdLine.customerInstructions("CUSTOMER_SPEC", cboUser_Def_Fld_2.Text.ToString.Trim())
        '++ID 3.24.2020 new message for COVID19
        cOrdLine.customerInstructions1(cboUser_Def_Fld_2.Text.ToString.Trim())



    End Sub

    '++ID 05012024 
    Private Sub cboWhyNotVegas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboWhyNotVegas.SelectedIndexChanged
        Try

            '    If cboWhyNotVegas.SelectedIndex = 0 Then Exit Sub
            '    If m_oOrder.Ordhead.Cus_No Is Nothing Then Exit Sub

            If txtOrd_No.Text <> "" And txtCus_No.Text <> "" Then

                Dim oTravReaTypeLateShip As cTraveler_Reason_Type_LateShipment
                oTravReaTypeLateShip = New cTraveler_Reason_Type_LateShipment

                oTravReaTypeLateShip = cboWhyNotVegas.SelectedItem

                If Not oTravReaTypeLateShip Is Nothing Then

                    ' If oTravReaTypeLateShip.ID <> 0 Then
                    If Not m_oOrder.Ordhead.Cus_No Is Nothing Then

                        'is declared as new in mGlobal for to be reusete in case if nobody will touch dropdown
                        'oTraveler_Reason_LateShipment = New cTraveler_Reason_LateShipment
                        If cboWhyNotVegas.SelectedIndex <> 0 Then

                            With oTraveler_Reason_LateShipment

                                .CUS_NO = m_oOrder.Ordhead.Cus_No
                                .REASON_TYPE_ID = CInt(cboWhyNotVegas.SelectedValue)
                                .OEIGUID = m_oOrder.Ordhead.Ord_GUID
                                .Save()
                            End With

                        Else

                            With oTraveler_Reason_LateShipment
                                .OEIGUID = m_oOrder.Ordhead.Ord_GUID
                                .Delete()
                            End With

                        End If


                    End If
                End If
            Else
                cboWhyNotVegas.SelectedIndex = 0
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub



    Private Sub cboreSSP_LostFocus(sender As Object, e As EventArgs) Handles cboreSSP.LostFocus
        Try

            If m_oOrder.Ordhead.intSSP <> cboreSSP.SelectedValue And m_oOrder.Ordhead.OEI_Ord_No <> "" Then
                Call Save()
            Else
                cboreSSP.SelectedValue = 0
            End If
        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try
    End Sub



End Class

Public Enum OrderNoChangeType
    NewOrder
    ExistingOrder
End Enum

