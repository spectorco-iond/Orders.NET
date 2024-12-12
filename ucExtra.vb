Option Strict Off
Option Explicit On

Friend Class ucExtra
	Inherits System.Windows.Forms.UserControl

    'Private m_oOrder As New cOrder
    Private blnLoading As Boolean = False

    Public Sub Fill() ' ByRef m_oOrder As cOrder)

        blnLoading = True

        'm_oOrder = m_oOrder

        If Trim(m_oOrder.Ordhead.Ord_No).Length <> 0 Then
            txtUser_Def_Fld_1.Text = m_oOrder.Ordhead.User_Def_Fld_1
            txtUser_Def_Fld_2.Text = m_oOrder.Ordhead.User_Def_Fld_2
            txtUser_Def_Fld_3.Text = m_oOrder.Ordhead.User_Def_Fld_3
            txtUser_Def_Fld_4.Text = m_oOrder.Ordhead.User_Def_Fld_4
            txtUser_Def_Fld_5.Text = m_oOrder.Ordhead.User_Def_Fld_5
        Else
            Reset() ' m_oOrder)
        End If

        blnLoading = False

    End Sub

    Public Sub Reset() ' ByRef pOrder As cOrder)

        blnLoading = True

        txtUser_Def_Fld_1.Text = ""
        txtUser_Def_Fld_2.Text = ""
        txtUser_Def_Fld_3.Text = ""
        txtUser_Def_Fld_4.Text = ""
        txtUser_Def_Fld_5.Text = ""

        blnLoading = False

    End Sub

    Public Sub Save()

        'm_oOrder.Ordhead.User_Def_Fld_1 = txtUser_Def_Fld_1.Text
        'm_oOrder.Ordhead.User_Def_Fld_2 = txtUser_Def_Fld_2.Text
        'm_oOrder.Ordhead.User_Def_Fld_3 = txtUser_Def_Fld_3.Text
        'm_oOrder.Ordhead.User_Def_Fld_4 = txtUser_Def_Fld_4.Text
        'm_oOrder.Ordhead.User_Def_Fld_5 = txtUser_Def_Fld_5.Text

    End Sub

    Public ReadOnly Property Order() As cOrder
        Get
            Order = m_oOrder
        End Get
    End Property

    Private Sub txtUser_Def_Fld_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)

        lblUser_Def_Fld_1.Focus()

    End Sub

    'Private Sub ucExtra_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

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
    '                oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.CreditInfo)
    '            Case Keys.D9
    '                'oFrmOrder.TabWindow.SelectTab(frmOrder.Tabs.Extra)
    '        End Select

    '    End If

    'End Sub

End Class

