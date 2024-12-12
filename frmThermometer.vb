Public Class frmThermometer

    Private m_Settings As cThermometer

    Private db As New cDBA()
    Private dtThermo As DataTable
    Private dtColors As DataTable


    Public Sub New(ByRef oSettings As cThermometer)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        m_Settings = oSettings

    End Sub

    Private Sub frmThermometer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Call CreateColumns()
            Call FillColumns()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub CreateColumns()

        Try
            dgvThermo.CausesValidation = False

            dgvThermo.Columns.Add(DGVTextBoxColumn("NEXT_AVAIL_DATE", "Date", 80))
            Me.Width = 132

            For Each oColumn As String In m_Settings.Methods
                dgvThermo.Columns.Add(DGVTextBoxColumn(oColumn.ToString.Trim.Replace(" ", "_"), oColumn.ToString.Trim, 60))
                dgvThermo.Columns.Add(DGVTextBoxColumn("COLOR_" & oColumn.ToString.Trim.Replace(" ", "_"), "Color " & oColumn.ToString.Trim, 60))
                dgvThermo.Columns("COLOR_" & oColumn.ToString.Trim.Replace(" ", "_")).Visible = False
                Me.Width += 60
            Next

            'dgvThermo.Columns(0).Width = 0
            'dgvThermo.RowHeadersWidth = 0

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub FillColumns()

        Dim iMethodColor As Integer

        Try

            Dim strSql As String
            strSql = _
            "SELECT		NEXT_AVAIL_DATE "

            For Each oMethod As String In m_Settings.Methods
                strSql &= _
                ", MAX(CASE METHOD_NAME WHEN '" & oMethod.ToString.Trim & "' THEN QTY ELSE 0 END) AS """ & oMethod.ToString.Trim.Replace(" ", "_") & """ " & _
                ", MAX(CASE METHOD_NAME WHEN '" & oMethod.ToString.Trim & "' THEN Color ELSE 0 END) AS ""COLOR_" & oMethod.ToString.Trim.Replace(" ", "_") & """ "
            Next

            strSql &= _
            "FROM EXACT_TRAVELER_THERMO_AVAIL_DATE " & _
            "GROUP BY	NEXT_AVAIL_DATE"

            dtThermo = db.DataTable(strSql)
            dgvThermo.DataSource = dtThermo

            Dim iRow As Integer = 0
            Dim iCol As Integer = 1

            If dtThermo.Rows.Count <> 0 Then
                For Each dtRow As DataRow In dtThermo.Rows
                    For Each oMethod As String In m_Settings.Methods
                        iMethodColor = dtRow.Item("COLOR_" & oMethod.ToString.Trim.Replace(" ", "_"))
                        With dgvThermo.Rows(iRow).Cells(iCol)
                            Select Case iMethodColor
                                Case 0
                                    If dgvThermo.Rows(iRow).Cells(iCol).Value = 0 Then
                                        .Style.BackColor = Color.Empty
                                    Else
                                        .Style.BackColor = Color.LightGreen
                                    End If
                                Case 1
                                    .Style.BackColor = Color.Orange ' Color.Orange
                                Case 2
                                    .Style.BackColor = Color.Yellow ' Color.Yellow
                                Case 3
                                    .Style.BackColor = Color.Cyan ' Color.Cyan
                                Case 4
                                    .Style.BackColor = Color.Red ' Color.Magenta
                                Case Else
                                    .Style.BackColor = Color.Empty

                            End Select
                        End With
                        iCol += 2
                    Next
                    iRow += 1
                    iCol = 1
                Next
            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Public Property Settings() As cThermometer
        Get
            Settings = m_Settings
        End Get
        Set(ByVal value As cThermometer)
            m_Settings = value
        End Set
    End Property

End Class