Option Strict Off
Option Explicit On

Friend Class ucImprint
    Inherits System.Windows.Forms.UserControl

    ' Private m_oOrder As New cOrder

    Private m_Imprint As New cImprint

    Private blnImprintLoading As Boolean = False
    Private m_ComboFill As Boolean = False

    Public Sub ImprintFill() ' ByRef pOrder As cOrder)

        Call ImprintComboFill()

        blnImprintLoading = True

        txtImprint.Text = m_Imprint.Imprint
        txtImprintLocation.Text = m_Imprint.Location
        txtNumImprints.Text = m_Imprint.Num_Impr_1
        txtImprintColor.Text = m_Imprint.Color
        txtPackaging.Text = m_Imprint.Packaging
        txtRefill.Text = m_Imprint.Refill
        'txtLaserSetup.text = m_Imprint.LaserSetup
        txtComments.Text = m_Imprint.Comments
        txtSpecialComments.Text = m_Imprint.Special_Comments
        txtIndustry.Text = m_Imprint.Industry

        blnImprintLoading = False

    End Sub

    Public Sub ImprintFill(ByRef poImprint As cImprint)

        m_Imprint = poImprint
        Call ImprintFill()

    End Sub

    Public Sub ImprintReset() ' ByRef pOrder As cOrder)

        m_Imprint.Reset()
        Call ImprintFill()

    End Sub

    Public Sub ImprintSave()

        ' We don't save OrdGUID and ItemGUID as they always stay the same here.
        m_Imprint.Imprint = txtImprint.Text
        m_Imprint.Location = txtImprintLocation.Text
        If IsNumeric(txtNumImprints.Text) Then
            m_Imprint.Num_Impr_1 = CInt(txtNumImprints.Text)
        Else
            m_Imprint.Num_Impr_1 = 0
        End If
        If IsNumeric(txtNumImprints.Text) Then
            m_Imprint.Num_Impr_1 = CInt(txtNumImprints.Text)
        Else
            m_Imprint.Num_Impr_1 = 0
        End If
        m_Imprint.Num = m_Imprint.Num_Impr_1
        m_Imprint.Color = txtImprintColor.Text
        m_Imprint.Packaging = txtPackaging.Text
        m_Imprint.Refill = txtRefill.Text
        ' m_Imprint.LaserSetup = txtLaserSetup.text
        m_Imprint.Comments = txtComments.Text
        m_Imprint.Special_Comments = txtSpecialComments.Text
        m_Imprint.Industry = txtIndustry.Text

    End Sub

    Private Sub ImprintComboFill()

        Dim strSql As String
        Dim conn As New cDBA
        Dim dtCombo As New DataTable

        If m_ComboFill Then Exit Sub

        Try
            strSql = "Select * from EXACT_TRAVELER_EXTRA_FIELDS_xRef order by FieldValue Asc "
            dtCombo = conn.DataTable(strSql)

            If dtCombo.Rows.Count <> 0 Then

                Dim drComboElement As DataRow

                For Each drComboElement In dtCombo.Rows

                    With drComboElement

                        Select Case .Item("Field")
                            Case "ImprintLocation"
                                cboImprintLocation.Items.Add(.Item("FieldValue"))
                            Case "ImprintColour"
                                cboImprintColor.Items.Add(.Item("FieldValue"))
                            Case "Industry"
                                cboIndustry.Items.Add(.Item("FieldValue"))
                            Case "Packaging"
                                cboPackaging.Items.Add(.Item("FieldValue"))
                            Case "Refill"
                                cboRefill.Items.Add(.Item("FieldValue"))
                        End Select

                    End With

                Next

                m_ComboFill = True

            End If

        Catch er As Exception
            MsgBox("Error in CImprint." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    'Public ReadOnly Property Order() As cOrder
    '    Get
    '        Order = m_oOrder
    '    End Get
    'End Property

    Private Sub cboImprintLocation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboImprintLocation.SelectedIndexChanged

        m_Imprint.AddLocation(cboImprintLocation.SelectedItem, txtImprintLocation.Text)
        txtImprintLocation.Text = m_Imprint.Location

    End Sub

    Private Sub cboImprintColor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboImprintColor.SelectedIndexChanged

        m_Imprint.AddColor(cboImprintColor.SelectedItem, txtImprintColor.Text)
        txtImprintColor.Text = m_Imprint.Color

    End Sub

    Private Sub cboIndustry_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboIndustry.SelectedIndexChanged

        m_Imprint.AddLocation(cboIndustry.SelectedItem, txtIndustry.Text)
        txtIndustry.Text = m_Imprint.Industry

    End Sub

    Private Sub cboPackaging_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPackaging.SelectedIndexChanged

        m_Imprint.AddPackaging(cboPackaging.SelectedItem, txtPackaging.Text)
        txtPackaging.Text = m_Imprint.Packaging

    End Sub

    Private Sub cboRefill_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRefill.SelectedIndexChanged

        m_Imprint.AddRefill(cboRefill.SelectedItem, txtRefill.Text)
        txtRefill.Text = m_Imprint.Refill

    End Sub

    Public Property Imprint(ByVal p_Imprint As cImprint)
        Get
            Imprint = m_Imprint
        End Get
        Set(ByVal value)
            m_Imprint = p_Imprint
        End Set
    End Property

    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Call Reset()

    End Sub

End Class
