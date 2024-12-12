Option Strict Off
Option Explicit On

Friend Class ucImprint_2
    Inherits System.Windows.Forms.UserControl

    ' Private m_oOrder As New cOrder
    Private blnLoading As Boolean = False
    Private m_ComboFill As Boolean = False

    Private m_Imprint As New cImprint

    Public Sub Fill() ' ByRef pOrder As cOrder)

        blnLoading = True

        txtImprint.Text = m_Imprint.Imprint
        txtImprintLocation.Text = m_Imprint.Location
        txtNumImprints.Text = m_Imprint.Num_Impr_1
        txtImprintColor.Text = m_Imprint.Color
        txtPackaging.Text = m_Imprint.Packaging
        txtRefill.Text = m_Imprint.Refill
        'txtLaserSetup.text = m_Imprint.LaserSetup
        txtComments.Text = m_Imprint.Comments
        txtSpecialComments.Text = m_Imprint.Special_Comments

        blnLoading = False

    End Sub

    Public Sub Reset() ' ByRef pOrder As cOrder)

        m_Imprint.Reset()
        Call Fill()

    End Sub

    Public Sub Save()

        ' We don't save OrdGUID and ItemGUID as they always stay the same here.
        m_Imprint.Imprint = txtImprint.Text
        m_Imprint.Location = txtImprintLocation.Text
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

    End Sub

    Private Sub ComboFill()

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

    Public ReadOnly Property Order() As cOrder
        Get
            Order = m_oOrder
        End Get
    End Property

    Private Sub cboImprintLocation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboImprintLocation.SelectedIndexChanged

        m_Imprint.AddLocation(txtImprintLocation.ToString, txtImprintLocation.Text)
        txtImprintLocation.Text = m_Imprint.Location

    End Sub

    Private Sub cboImprintColor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboImprintColor.SelectedIndexChanged

        m_Imprint.AddColor(txtImprintColor.ToString, txtImprintColor.Text)
        txtImprintColor.Text = m_Imprint.Color

    End Sub

    Private Sub cboPackaging_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPackaging.SelectedIndexChanged

        m_Imprint.AddPackaging(txtPackaging.ToString, txtPackaging.Text)
        txtPackaging.Text = m_Imprint.Packaging

    End Sub

    Private Sub cboRefill_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRefill.SelectedIndexChanged

        m_Imprint.AddRefill(txtRefill.ToString, txtRefill.Text)
        txtRefill.Text = m_Imprint.Refill

    End Sub

    Public Property Imprint(ByVal p_Imprint As cImprint)
        Get
            Imprint = m_Imprint
            'Imprint = New cImprint

            'Imprint.Color = m_Imprint.Color
            'Imprint.Comments = m_Imprint.Comments
            'Imprint.Imprint = m_Imprint.Imprint
            'Imprint.Item_Guid = m_Imprint.Item_Guid
            'Imprint.LaserSetup = m_Imprint.LaserSetup
            'Imprint.Location = m_Imprint.Location
            'Imprint.Num = m_Imprint.Location
            'Imprint.Ord_Guid = m_Imprint.Ord_Guid
            'Imprint.Packaging = m_Imprint.Packaging
            'Imprint.Refill = m_Imprint.Refill

        End Get
        Set(ByVal value)
            m_Imprint = p_Imprint
            'm_Imprint = New cImprint
            'm_Imprint.Color = p_Imprint.Color
            'm_Imprint.Comments = p_Imprint.Comments
            'm_Imprint.Imprint = p_Imprint.Imprint
            'm_Imprint.Item_Guid = p_Imprint.Item_Guid
            'm_Imprint.LaserSetup = p_Imprint.LaserSetup
            'm_Imprint.Location = p_Imprint.Location
            'm_Imprint.Num = p_Imprint.Location
            'm_Imprint.Ord_Guid = p_Imprint.Ord_Guid
            'm_Imprint.Packaging = p_Imprint.Packaging
            'm_Imprint.Refill = p_Imprint.Refill

        End Set
    End Property

    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click

        Call Reset()

    End Sub

    Private Sub cmdGetRepeatData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetRepeatData.Click

    End Sub
End Class
