Public Class frmKitSelection

    Dim m_Kit As CKit

    Private Enum Columns
        FirstCol = 0
        Item_No
        Item_Desc_1
        Item_Desc_2
        Item_Price
        Comp_Item_No
        Comp_Item_Price
        Comp_Item_Desc_1
        Comp_Item_Desc_2
        Comp_Qty_Per_Par
        Comp_Qty_Available
    End Enum

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Me.Close()

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

        ' Do nothing for now

    End Sub
    
    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click

        ' Do nothing for now

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        ' Do nothing for now

    End Sub

    Public Property Kit() As cKit
        Get
            Kit = m_Kit
        End Get
        Set(ByVal value As cKit)
            m_Kit = value
        End Set
    End Property

    Private Sub frmKitSelection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            If m_Kit Is Nothing Then Exit Sub

            For iPos As Integer = (dgvItems.Rows.Count - 1) To 0 Step -1
                dgvItems.Rows.RemoveAt(iPos)
            Next

            txtItem_No.Text = m_Kit.Item_No
            txtDescription_1.Text = m_Kit.Item_Desc_1
            txtDescription_2.Text = m_Kit.Item_Desc_2
            txtUnit_Price.Text = m_Kit.Item_Price

            dgvItems.Columns.Add(DGVTextBoxColumn("FirstCol", "FirstCol", 0))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_No", "Line", 110))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Desc_1", "Line", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Desc_2", "Line", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Item_Price", "Line", 100))
            dgvItems.Columns.Add(DGVTextBoxColumn("Comp_Item_No", "Component Item", 115))
            dgvItems.Columns.Add(DGVTextBoxColumn("Comp_Item_Price", "Unit price", 80))
            dgvItems.Columns.Add(DGVTextBoxColumn("Comp_Item_Desc_1", "Description", 230))
            dgvItems.Columns.Add(DGVTextBoxColumn("Comp_Item_Desc_2", "Component Description 2", 220))
            dgvItems.Columns.Add(DGVTextBoxColumn("Comp_Qty_Per_Par", "Quantity per kit", 80))
            dgvItems.Columns.Add(DGVTextBoxColumn("Comp_Qty_Available", "Quantity available", 90))

            Dim dt As DataTable
            Dim db As New cDBA

            dt = db.DataTable(m_Kit.Sql)

            dgvItems.DataSource = dt

            dgvItems.AllowUserToAddRows = False
            dgvItems.AllowUserToOrderColumns = False

            For lPos As Integer = 0 To dgvItems.Columns.Count - 1
                dgvItems.Columns(lPos).SortMode = DataGridViewColumnSortMode.NotSortable
            Next lPos

            dgvItems.CausesValidation = True

            Dim oCellStyle As New DataGridViewCellStyle()
            oCellStyle = New DataGridViewCellStyle()
            oCellStyle.Format = "##,##0.00"
            oCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvItems.Columns(Columns.Comp_Qty_Per_Par).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Comp_Item_Price).DefaultCellStyle = oCellStyle
            dgvItems.Columns(Columns.Comp_Qty_Available).DefaultCellStyle = oCellStyle

            dgvItems.Columns(Columns.FirstCol).Frozen = True
            dgvItems.Columns(Columns.Comp_Item_No).Frozen = True

            dgvItems.Columns(Columns.FirstCol).Visible = False
            dgvItems.Columns(Columns.Item_No).Visible = False
            dgvItems.Columns(Columns.Item_Desc_1).Visible = False
            dgvItems.Columns(Columns.Item_Desc_2).Visible = False
            dgvItems.Columns(Columns.Item_Price).Visible = False

            dgvItems.CurrentCell = dgvItems.Rows(0).Cells(Columns.Comp_Item_No)

            cmdClose.Focus()

        Catch er As Exception
            MsgBox("Error in " & Me.Name & "." & New Diagnostics.StackFrame(0).GetMethod.Name & vbCrLf & er.Message)
        End Try

    End Sub

    Private Function DGVTextBoxColumn(ByVal pstrName As String, ByVal pstrHeaderText As String, Optional ByVal plWidth As Long = 100) As DataGridViewTextBoxColumn

        DGVTextBoxColumn = New DataGridViewTextBoxColumn

        DGVTextBoxColumn.HeaderText = pstrHeaderText
        DGVTextBoxColumn.DataPropertyName = pstrName
        DGVTextBoxColumn.Name = pstrName

        DGVTextBoxColumn.Visible = (Width <> 0)
        DGVTextBoxColumn.Width = plWidth

    End Function

    Private Sub dgvItems_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvItems.KeyDown

        Select Case e.KeyData
            Case Keys.Escape
                Me.Close()

        End Select

    End Sub

    Private Sub TextBoxGotFocus() Handles txtDescription_1.GotFocus, txtDescription_2.GotFocus, txtItem_No.GotFocus, txtUnit_Price.GotFocus

        dgvItems.Focus()

    End Sub

    Private Sub Escape_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
                cmdClose.KeyDown, dgvItems.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose.PerformClick()
        End Select

    End Sub

End Class