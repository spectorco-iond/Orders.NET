Option Strict Off
Option Explicit On

Friend Class ucCreditInfo
    Inherits System.Windows.Forms.UserControl

    'Private m_oOrder As New cOrder
    Private blnLoading As Boolean = False

    Public Sub Fill() ' ByRef m_oOrder As cOrder)

        blnLoading = True

        'm_oOrder = m_oOrder

        If Trim(m_oOrder.Ordhead.OEI_Ord_No).Length <> 0 Then
            lblAmt_Cus_Balance.Text = String.Format("{0:n2}", m_oOrder.CreditInfo.Amt_Cus_Balance)
            lblAmt_Cur_Orders.Text = String.Format("{0:n2}", m_oOrder.CreditInfo.Amt_Cur_Orders)
            lblAmt_Open_Orders.Text = String.Format("{0:n2}", m_oOrder.CreditInfo.Amt_Open_Orders)
            lblAmt_Orders_On_Hold.Text = String.Format("{0:n2}", m_oOrder.CreditInfo.Amt_Orders_On_Hold)
            lblAmt_Open_Invoices.Text = String.Format("{0:n2}", m_oOrder.CreditInfo.Amt_Open_Invoices)
            lblAmt_Credit_Limit.Text = String.Format("{0:n0}", m_oOrder.CreditInfo.Amt_Credit_Limit)
            lblAmt_Avail_Credit.Text = String.Format("{0:n2}", m_oOrder.CreditInfo.Amt_Avail_Credit)
            lblDays_Over_Terms.Text = m_oOrder.CreditInfo.Days_Over_Terms
            lblAmt_Over_Terms.Text = String.Format("{0:n2}", m_oOrder.CreditInfo.Amt_Over_Terms)
            lblCredit_Hold.Text = m_oOrder.CreditInfo.Credit_Hold
            lblCredit_Rate.Text = m_oOrder.CreditInfo.Credit_Rate
            lblAcc_Start_Dt.Text = m_oOrder.CreditInfo.Acc_Start_Dt
            lblComment1.Text = m_oOrder.CreditInfo.Comment1
            lblComment2.Text = m_oOrder.CreditInfo.Comment2
        Else
            Reset() ' m_oOrder)
        End If

        blnLoading = False

    End Sub

    Public Sub Reset() ' ByRef pOrder As cOrder)

        lblAmt_Cus_Balance.Text = ""
        lblAmt_Cur_Orders.Text = ""
        lblAmt_Open_Orders.Text = ""
        lblAmt_Orders_On_Hold.Text = ""
        lblAmt_Open_Invoices.Text = ""
        lblAmt_Credit_Limit.Text = ""
        lblAmt_Avail_Credit.Text = ""
        lblDays_Over_Terms.Text = ""
        lblAmt_Over_Terms.Text = ""
        lblCredit_Hold.Text = ""
        lblCredit_Rate.Text = ""
        lblAcc_Start_Dt.Text = ""
        lblComment1.Text = ""
        lblComment2.Text = ""

    End Sub

    Public Sub Save()

    End Sub

    Public ReadOnly Property Order() As cOrder
        Get
            Order = m_oOrder
        End Get
    End Property

    'Private Sub ucCreditInfo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

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
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Salesperson)
    '            Case Keys.D8
    '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
    '            Case Keys.D9
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
    '        End Select

    '    End If

    'End Sub

End Class