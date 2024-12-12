Public Class frmItemETA

    Private m_Item_No As String

    Public Sub New(ByVal pstrItem_No As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        m_Item_No = pstrItem_No

        Call DataLoad()

    End Sub

    Private Sub DataLoad()

        Try

            Dim dt As DataTable
            Dim db As New cDBA

            Dim strSql As String = _
            "SELECT     I.Item_No AS 'Item No', ISNULL(Q.qtyremaining, 0) AS 'Qty Remaining', " & _
            "           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as 'Promise Date', " & _
            "           Ord_No AS 'Ord No', Vend_No AS 'Vend No', CONVERT(CHAR(12), Request_Dt, 103) AS 'Request Date', PO_Type AS 'PO Type' " & _
            "FROM       IMItmIdx_Sql I WITH (Nolock) " & _
            "LEFT JOIN  Bankers_Cust_poordlin_Qtyremaining_ETA Q ON I.Item_No = Q.Item_No " & _
            "WHERE      I.Item_No = '" & Trim(m_Item_No) & "' " & _
            "ORDER BY   Q.Promise_Dt "

            'ID 02.17.2023
            If m_oOrder.Ordhead.Mfg_Loc = "US1" Then
                strSql =
            "SELECT     I.Item_No AS 'Item No', ISNULL(Q.qtyremaining, 0) AS 'Qty Remaining', " &
            "           CASE WHEN Promise_Dt IS NULL THEN '' ELSE CONVERT(CHAR(12), Promise_Dt, 103) END as 'Promise Date', " &
            "           Ord_No AS 'Ord No', Vend_No AS 'Vend No', CONVERT(CHAR(12), Request_Dt, 103) AS 'Request Date', PO_Type AS 'PO Type' " &
            "FROM       [200].dbo.IMItmIdx_Sql I WITH (Nolock) " &
            "LEFT JOIN  [200].dboBankers_Cust_poordlin_Qtyremaining_ETA_US Q ON I.Item_No = Q.Item_No " &
            "WHERE      I.Item_No = '" & Trim(m_Item_No) & "' " &
            "ORDER BY   Q.Promise_Dt "
            End If


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                txtItem_No.Text = dt.Rows(0).Item("Item No").ToString

                dgvLocDetails.DataSource = dt

                dgvLocDetails.RowHeadersWidth = 20

                dgvLocDetails.Columns(Columns.Item_No).Width = 125
                dgvLocDetails.Columns(Columns.QtyRemaining).Width = 80
                dgvLocDetails.Columns(Columns.Promise_Date).Width = 80
                dgvLocDetails.Columns(Columns.Ord_No).Width = 80
                dgvLocDetails.Columns(Columns.Vend_No).Width = 100
                dgvLocDetails.Columns(Columns.Request_Dt).Width = 80
                dgvLocDetails.Columns(Columns.Po_Type).Width = 80

                dgvLocDetails.Columns(Columns.Item_No).Visible = False

                Dim oCellStyle As New DataGridViewCellStyle()
                oCellStyle.Format = "##,##0.00"
                oCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                dgvLocDetails.Columns(Columns.QtyRemaining).DefaultCellStyle = oCellStyle

            End If

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Me.Close()

    End Sub

    Private Enum Columns

        Item_No
        QtyRemaining
        Promise_Date
        Ord_No
        Vend_No
        Request_Dt
        Po_Type

    End Enum

End Class