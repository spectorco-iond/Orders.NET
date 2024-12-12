Public Class frmItemInfo

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

            Dim strSql As String
            'strSql = "SELECT * FROM DBO.OEI_Item_Price_Breaks('" & m_Curr_Cd & "', '" & m_Cus_No & "', '" & m_Item_No & "', '" & m_Start_Dt & "', '" & m_End_Dt & "')"


            If m_oOrder.Ordhead.Mfg_Loc <> "US1" Then
                strSql = "EXEC DBO.OEI_Item_Location_Detail_Form '" & m_Item_No & "' "
            Else
                'ID 02.17.2023
                strSql = "EXEC [200].DBO.OEI_Item_Location_Detail_Form_US '" & m_Item_No & "' "
            End If


            dt = db.DataTable(strSql)

            If dt.Rows.Count <> 0 Then

                txtItem_No.Text = dt.Rows(0).Item("Item_No").ToString
                txtItem_Desc_1.Text = dt.Rows(0).Item("Item_Desc_1").ToString
                txtItem_Desc_2.Text = dt.Rows(0).Item("Item_Desc_2").ToString

                txtSum_Qty_BkOrd.Text = dt.Rows(0).Item("Sum_Qty_BkOrd").ToString
                txtSum_Qty_On_Ord.Text = dt.Rows(0).Item("Sum_Qty_On_Ord").ToString
                txtSum_Qty_On_Ord_PO.Text = dt.Rows(0).Item("Sum_Qty_On_Ord_PO").ToString
                txtSum_Qty_In_Transit.Text = dt.Rows(0).Item("Sum_Qty_In_Transit").ToString
                txtMax_Reorder_Lvl.Text = dt.Rows(0).Item("Max_Reorder_Lvl").ToString

                dgvLocDetails.DataSource = dt

                dgvLocDetails.Columns(Columns.Item_No).Visible = False
                dgvLocDetails.Columns(Columns.Item_Desc_1).Visible = False
                dgvLocDetails.Columns(Columns.Item_Desc_2).Visible = False
                dgvLocDetails.Columns(Columns.Sum_Qty_BkOrd).Visible = False
                dgvLocDetails.Columns(Columns.Sum_Qty_On_Ord).Visible = False
                dgvLocDetails.Columns(Columns.Sum_Qty_On_Ord_PO).Visible = False
                dgvLocDetails.Columns(Columns.Sum_Qty_In_Transit).Visible = False
                dgvLocDetails.Columns(Columns.Max_Reorder_Lvl).Visible = False

                dgvLocDetails.RowHeadersWidth = 20

                dgvLocDetails.Columns(Columns.Loc).Width = 50
                dgvLocDetails.Columns(Columns.Loc_Desc).Width = 120
                dgvLocDetails.Columns(Columns.Qty_Allocated).Width = 120
                dgvLocDetails.Columns(Columns.Qty_On_Hand).Width = 120
                dgvLocDetails.Columns(Columns.Qty_Available).Width = 120

                Dim oCellStyle As New DataGridViewCellStyle()
                oCellStyle.Format = "##,##0.00"
                oCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

                dgvLocDetails.Columns(Columns.Qty_Allocated).DefaultCellStyle = oCellStyle
                dgvLocDetails.Columns(Columns.Qty_On_Hand).DefaultCellStyle = oCellStyle
                dgvLocDetails.Columns(Columns.Qty_Available).DefaultCellStyle = oCellStyle

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
        Item_Desc_1
        Item_Desc_2
        Loc
        Loc_Desc
        Qty_On_Hand
        Qty_Allocated
        Qty_Available
        Sum_Qty_BkOrd
        Sum_Qty_On_Ord
        Sum_Qty_On_Ord_PO
        Sum_Qty_In_Transit
        Max_Reorder_Lvl

    End Enum

End Class