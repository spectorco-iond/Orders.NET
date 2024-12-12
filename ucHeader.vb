Option Strict Off
Option Explicit On

Friend Class ucHeader
    Inherits System.Windows.Forms.UserControl

    'Private m_oOrder As New cOrder
    Private blnLoading As Boolean = False

    Public Sub Fill() ' ByRef m_oOrder As cOrder)

        blnLoading = True

        'm_oOrder = m_oOrder

        If Trim(m_oOrder.Ordhead.Ord_No).Length <> 0 Then
            txtOrd_Type.Text = m_oOrder.Ordhead.Ord_Type
            txtOrd_No.Text = m_oOrder.Ordhead.Ord_No
            txtOrd_Dt.Text = m_oOrder.Ordhead.Ord_Dt
            txtApply_To_No.Text = m_oOrder.Ordhead.Apply_To_No
            txtCus_No.Text = m_oOrder.Ordhead.Cus_No
            txtCustomerName.Text = m_oOrder.Ordhead.Bill_To_Name

            txtOE_PO_No.Text = m_oOrder.Ordhead.Oe_Po_No
            txtShipping_Dt.Text = m_oOrder.Ordhead.Shipping_Dt
            txtCus_Alt_Adr_Cd.Text = m_oOrder.Ordhead.Cus_Alt_Adr_Cd
            txtShip_To_Name.Text = m_oOrder.Ordhead.Ship_To_Name

            txtShip_Via_Cd.Text = m_oOrder.Ordhead.Ship_Via_Cd
            lblShip_Via_Desc.Text = m_oOrder.Descriptions.ShipVia
            txtAR_Terms_Cd.Text = m_oOrder.Ordhead.Ar_Terms_Cd
            lblAR_Terms_Desc.Text = m_oOrder.Descriptions.Terms
            txtMfg_Loc.Text = m_oOrder.Ordhead.Mfg_Loc
            lblMfg_Loc_Desc.Text = m_oOrder.Descriptions.Location
            txtJob_No.Text = m_oOrder.Ordhead.Job_No

            lblStatus.Text = m_oOrder.Ordhead.Status
            lblStatus_Desc.Text = m_oOrder.Descriptions.Status
            txtDiscount_Pct.Text = m_oOrder.Ordhead.Discount_Pct
            'txtPart_Posted_FG.Text = m_oOrder.Ordhead.Part_Posted_Fg
            cboPart_Posted_FG.Text = "N=None"

            txtHold_FG.Text = m_oOrder.Ordhead.Hold_Fg
            txtForm_No.Text = m_oOrder.Ordhead.Form_No
            txtAward_Probability.Text = m_oOrder.Ordhead.Award_Probability

            txtLost_Sale_Cd.Text = m_oOrder.Ordhead.Lost_Sale_Cd
            txtCurr_Cd.Text = m_oOrder.Ordhead.Curr_Cd
            txtCurr_Trx_Rt.Text = m_oOrder.Ordhead.Curr_Trx_Rt
            'txtDeter_Rate_By.Text = m_oOrder.Ordhead.Deter_Rate_By
            cboDeter_Rate_By.Text = "I=Invoice Date"

            txtProfit_Center.Text = m_oOrder.Ordhead.Profit_Center
            txtDept.Text = m_oOrder.Ordhead.Dept
            'txtDept_Desc.Text = m_oOrder.Ordhead.Dept_Desc
            txtShip_Instruction_1.Text = m_oOrder.Ordhead.Ship_Instruction_1
            txtShip_Instruction_2.Text = m_oOrder.Ordhead.Ship_Instruction_2
            txtOrd_Guid.Text = m_oOrder.Ordhead.Ord_GUID
        Else
            Reset() ' m_oOrder)
        End If

        blnLoading = False

    End Sub

    Public Sub Reset() ' ByRef pOrder As cOrder)

        blnLoading = True

        txtOrd_Type.Text = ""
        txtOrd_No.Text = ""
        txtOrd_Dt.Text = ""
        txtApply_To_No.Text = ""
        txtCus_No.Text = ""
        txtCustomerName.Text = ""

        txtOE_PO_No.Text = ""
        txtShipping_Dt.Text = ""
        txtCus_Alt_Adr_Cd.Text = ""
        txtShip_To_Name.Text = ""

        txtShip_Via_Cd.Text = ""
        lblShip_Via_Desc.Text = ""
        txtAR_Terms_Cd.Text = ""
        lblAR_Terms_Desc.Text = ""
        txtMfg_Loc.Text = ""
        lblMfg_Loc_Desc.Text = ""
        txtJob_No.Text = ""

        lblStatus.Text = ""
        lblStatus_Desc.Text = ""
        txtDiscount_Pct.Text = ""
        'txtPart_Posted_FG.Text = "N"
        cboPart_Posted_FG.Text = "N=None"

        txtHold_FG.Text = ""
        txtForm_No.Text = ""
        txtAward_Probability.Text = ""

        txtLost_Sale_Cd.Text = ""
        txtCurr_Cd.Text = ""
        txtCurr_Trx_Rt.Text = ""
        'txtDeter_Rate_By.Text = "I"
        cboDeter_Rate_By.Text = "I=Invoice Date"

        txtProfit_Center.Text = ""
        txtCostCenter.Text = ""
        txtDept.Text = ""
        txtDept2.Text = ""
        txtDept3.Text = ""
        txtDept_Desc.Text = ""
        txtShip_Instruction_1.Text = ""
        txtShip_Instruction_2.Text = ""
        txtOrd_Guid.Text = String.Empty

        blnLoading = False

    End Sub

    Public Sub Save()

        Exit Sub

        ' Do nothing here for now
        m_oOrder.Ordhead.Ord_Type = txtOrd_Type.Text
        m_oOrder.Ordhead.Ord_No = txtOrd_No.Text
        m_oOrder.Ordhead.Ord_Dt = txtOrd_Dt.Text
        m_oOrder.Ordhead.Apply_To_No = txtApply_To_No.Text
        m_oOrder.Ordhead.Cus_No = txtCus_No.Text
        'txtCustomerName = txtCustomerName

        m_oOrder.Ordhead.Oe_Po_No = txtOE_PO_No.Text
        m_oOrder.Ordhead.Shipping_Dt = txtShipping_Dt.Text
        m_oOrder.Ordhead.Cus_Alt_Adr_Cd = txtCus_Alt_Adr_Cd.Text
        m_oOrder.Ordhead.Ship_To_Name = txtShip_To_Name.Text

        m_oOrder.Ordhead.Ship_Via_Cd = txtShip_Via_Cd.Text
        'lblShip_Via_Desc = txtShip_Via_Desc
        m_oOrder.Ordhead.Ar_Terms_Cd = txtAR_Terms_Cd.Text
        'lblAR_Terms_Desc = txtAR_Terms_Desc
        m_oOrder.Ordhead.Mfg_Loc = txtMfg_Loc.Text
        'lblMfg_Loc_Desc = txtMfg_Loc_Desc
        m_oOrder.Ordhead.Job_No = txtJob_No.Text

        m_oOrder.Ordhead.Status = lblStatus.Text
        'lblStatus_Desc = txtStatus_Desc
        m_oOrder.Ordhead.Discount_Pct = txtDiscount_Pct.Text
        m_oOrder.Ordhead.Part_Posted_Fg = txtPart_Posted_FG.Text

        m_oOrder.Ordhead.Hold_Fg = txtHold_FG.Text
        m_oOrder.Ordhead.Form_No = txtForm_No.Text
        m_oOrder.Ordhead.Award_Probability = txtAward_Probability.Text

        m_oOrder.Ordhead.Lost_Sale_Cd = txtLost_Sale_Cd.Text
        m_oOrder.Ordhead.Curr_Cd = txtCurr_Cd.Text
        m_oOrder.Ordhead.Curr_Trx_Rt = txtCurr_Trx_Rt.Text
        m_oOrder.Ordhead.Deter_Rate_By = Mid(txtDeter_Rate_By.Text, 1, 1)

        m_oOrder.Ordhead.Profit_Center = txtProfit_Center.Text
        m_oOrder.Ordhead.Dept = txtDept.Text
        'txtDept_Desc = txtDept_Desc
        m_oOrder.Ordhead.Ship_Instruction_1 = txtShip_Instruction_1.Text
        m_oOrder.Ordhead.Ship_Instruction_2 = txtShip_Instruction_2.Text

    End Sub

    Public ReadOnly Property Order() As cOrder
        Get
            Order = m_oOrder
        End Get
    End Property

    Private Sub cboPart_Posted_FG_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPart_Posted_FG.Click

    End Sub

    Private Sub cboPart_Posted_FG_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPart_Posted_FG.SelectedIndexChanged
        txtPart_Posted_FG.Text = Mid(cboPart_Posted_FG.Text, 1, 1)
    End Sub

    Private Sub cboDeter_Rate_By_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDeter_Rate_By.SelectedIndexChanged
        txtDeter_Rate_By.Text = IIf(cboDeter_Rate_By.Text = "I=Invoice Date", "I", "")
    End Sub

    Private Sub txtPart_Posted_FG_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPart_Posted_FG.TextChanged
        cboPart_Posted_FG.Text = IIf(txtPart_Posted_FG.Text = "C", "C=Col", IIf(txtPart_Posted_FG.Text = "N", "N=None", IIf(txtPart_Posted_FG.Text = "P", "P=Ppd", "")))
    End Sub

    Private Sub txtDeter_Rate_By_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDeter_Rate_By.TextChanged
        cboDeter_Rate_By.Text = IIf(txtDeter_Rate_By.Text = "I", "I=Invoice Date", "")
    End Sub

    Private Sub Controls_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDiscount_Pct.GotFocus, txtHold_FG.GotFocus, _
        txtAward_Probability.GotFocus, txtCurr_Trx_Rt.GotFocus, txtUnknown.GotFocus, txtCostCenter.GotFocus, txtProfit_Center.GotFocus, _
        txtDept.GotFocus, txtDept2.GotFocus, txtDept3.GotFocus, txtShip_Instruction_1.GotFocus, txtShip_Instruction_2.GotFocus, _
        txtJob_No.GotFocus, txtMfg_Loc.GotFocus, txtAR_Terms_Cd.GotFocus, txtShip_Via_Cd.GotFocus, txtShip_To_Name.GotFocus, _
        txtCus_Alt_Adr_Cd.GotFocus, txtShipping_Dt.GotFocus, txtOE_PO_No.GotFocus, txtCustomerName.GotFocus, txtCus_No.GotFocus, _
        txtApply_To_No.GotFocus, txtOrd_Dt.GotFocus, txtOrd_No.GotFocus, txtOrd_Type.GotFocus, txtForm_No.GotFocus

        lblOrd_Type.Focus()

    End Sub

    Private Sub txtShip_Via_Cd_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtShip_Via_Cd.TextChanged
        m_oOrder.Descriptions.ShipVia = m_oOrder.Ordhead.Get_Ship_Via_Cd_Description(m_oOrder.Ordhead.Ship_Via_Cd)
        lblShip_Via_Desc.Text = m_oOrder.Descriptions.ShipVia
    End Sub

    Private Sub txtAR_Terms_Cd_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAR_Terms_Cd.TextChanged
        m_oOrder.Descriptions.Terms = m_oOrder.Ordhead.Get_Ar_Terms_Cd_Description(m_oOrder.Ordhead.Ar_Terms_Cd)
        lblAR_Terms_Desc.Text = m_oOrder.Descriptions.Terms
    End Sub

    Private Sub txtMfg_Loc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMfg_Loc.TextChanged
        m_oOrder.Descriptions.Location = m_oOrder.Ordhead.Get_Loc_Description(m_oOrder.Ordhead.Mfg_Loc)
        lblMfg_Loc_Desc.Text = m_oOrder.Descriptions.Location
    End Sub

    'Private Sub ucHeader_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

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
    '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.HeaderAll)
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

End Class
