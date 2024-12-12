<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucOrderTotal
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblShip_Amt = New System.Windows.Forms.Label
        Me.lblOrig_Curr_Amt = New System.Windows.Forms.Label
        Me.lblOrder_Amt = New System.Windows.Forms.Label
        Me.lblTo_Curr_Cd_Desc = New System.Windows.Forms.Label
        Me.lblFrom_Curr_Cd_Desc = New System.Windows.Forms.Label
        Me.Label38 = New System.Windows.Forms.Label
        Me.Label37 = New System.Windows.Forms.Label
        Me.lblShip_Qty = New System.Windows.Forms.Label
        Me.lblConv_Curr_Amt = New System.Windows.Forms.Label
        Me.lblOrder_Qty = New System.Windows.Forms.Label
        Me.lblCurr_Trx_Rt = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.lblOrder_Item_Count = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblShip_Amt
        '
        Me.lblShip_Amt.BackColor = System.Drawing.SystemColors.Control
        Me.lblShip_Amt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblShip_Amt.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblShip_Amt.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblShip_Amt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShip_Amt.Location = New System.Drawing.Point(474, 9)
        Me.lblShip_Amt.Name = "lblShip_Amt"
        Me.lblShip_Amt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblShip_Amt.Size = New System.Drawing.Size(120, 20)
        Me.lblShip_Amt.TabIndex = 4
        Me.lblShip_Amt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOrig_Curr_Amt
        '
        Me.lblOrig_Curr_Amt.BackColor = System.Drawing.SystemColors.Control
        Me.lblOrig_Curr_Amt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOrig_Curr_Amt.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOrig_Curr_Amt.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOrig_Curr_Amt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrig_Curr_Amt.Location = New System.Drawing.Point(598, 9)
        Me.lblOrig_Curr_Amt.Name = "lblOrig_Curr_Amt"
        Me.lblOrig_Curr_Amt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOrig_Curr_Amt.Size = New System.Drawing.Size(120, 20)
        Me.lblOrig_Curr_Amt.TabIndex = 5
        Me.lblOrig_Curr_Amt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOrder_Amt
        '
        Me.lblOrder_Amt.BackColor = System.Drawing.SystemColors.Control
        Me.lblOrder_Amt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOrder_Amt.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOrder_Amt.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOrder_Amt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrder_Amt.Location = New System.Drawing.Point(350, 9)
        Me.lblOrder_Amt.Name = "lblOrder_Amt"
        Me.lblOrder_Amt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOrder_Amt.Size = New System.Drawing.Size(120, 20)
        Me.lblOrder_Amt.TabIndex = 3
        Me.lblOrder_Amt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTo_Curr_Cd_Desc
        '
        Me.lblTo_Curr_Cd_Desc.BackColor = System.Drawing.SystemColors.Control
        Me.lblTo_Curr_Cd_Desc.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTo_Curr_Cd_Desc.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTo_Curr_Cd_Desc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTo_Curr_Cd_Desc.Location = New System.Drawing.Point(724, 33)
        Me.lblTo_Curr_Cd_Desc.Name = "lblTo_Curr_Cd_Desc"
        Me.lblTo_Curr_Cd_Desc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTo_Curr_Cd_Desc.Size = New System.Drawing.Size(133, 20)
        Me.lblTo_Curr_Cd_Desc.TabIndex = 12
        Me.lblTo_Curr_Cd_Desc.Text = "Canadian Dollars"
        '
        'lblFrom_Curr_Cd_Desc
        '
        Me.lblFrom_Curr_Cd_Desc.BackColor = System.Drawing.SystemColors.Control
        Me.lblFrom_Curr_Cd_Desc.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFrom_Curr_Cd_Desc.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFrom_Curr_Cd_Desc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFrom_Curr_Cd_Desc.Location = New System.Drawing.Point(724, 10)
        Me.lblFrom_Curr_Cd_Desc.Name = "lblFrom_Curr_Cd_Desc"
        Me.lblFrom_Curr_Cd_Desc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFrom_Curr_Cd_Desc.Size = New System.Drawing.Size(102, 20)
        Me.lblFrom_Curr_Cd_Desc.TabIndex = 6
        Me.lblFrom_Curr_Cd_Desc.Text = "US Dollars"
        '
        'Label38
        '
        Me.Label38.BackColor = System.Drawing.SystemColors.Control
        Me.Label38.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label38.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label38.Location = New System.Drawing.Point(257, 10)
        Me.Label38.Name = "Label38"
        Me.Label38.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label38.Size = New System.Drawing.Size(101, 20)
        Me.Label38.TabIndex = 2
        Me.Label38.Text = "Sale Amount:"
        '
        'Label37
        '
        Me.Label37.BackColor = System.Drawing.SystemColors.Control
        Me.Label37.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label37.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label37.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label37.Location = New System.Drawing.Point(6, 33)
        Me.Label37.Name = "Label37"
        Me.Label37.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label37.Size = New System.Drawing.Size(113, 20)
        Me.Label37.TabIndex = 7
        Me.Label37.Text = "Currency rate:"
        '
        'lblShip_Qty
        '
        Me.lblShip_Qty.BackColor = System.Drawing.SystemColors.Control
        Me.lblShip_Qty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblShip_Qty.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblShip_Qty.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblShip_Qty.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblShip_Qty.Location = New System.Drawing.Point(474, 32)
        Me.lblShip_Qty.Name = "lblShip_Qty"
        Me.lblShip_Qty.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblShip_Qty.Size = New System.Drawing.Size(120, 20)
        Me.lblShip_Qty.TabIndex = 10
        Me.lblShip_Qty.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblConv_Curr_Amt
        '
        Me.lblConv_Curr_Amt.BackColor = System.Drawing.SystemColors.Control
        Me.lblConv_Curr_Amt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblConv_Curr_Amt.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblConv_Curr_Amt.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConv_Curr_Amt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblConv_Curr_Amt.Location = New System.Drawing.Point(598, 32)
        Me.lblConv_Curr_Amt.Name = "lblConv_Curr_Amt"
        Me.lblConv_Curr_Amt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblConv_Curr_Amt.Size = New System.Drawing.Size(120, 20)
        Me.lblConv_Curr_Amt.TabIndex = 11
        Me.lblConv_Curr_Amt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOrder_Qty
        '
        Me.lblOrder_Qty.BackColor = System.Drawing.SystemColors.Control
        Me.lblOrder_Qty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOrder_Qty.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOrder_Qty.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOrder_Qty.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrder_Qty.Location = New System.Drawing.Point(350, 32)
        Me.lblOrder_Qty.Name = "lblOrder_Qty"
        Me.lblOrder_Qty.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOrder_Qty.Size = New System.Drawing.Size(120, 20)
        Me.lblOrder_Qty.TabIndex = 9
        Me.lblOrder_Qty.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCurr_Trx_Rt
        '
        Me.lblCurr_Trx_Rt.BackColor = System.Drawing.SystemColors.Control
        Me.lblCurr_Trx_Rt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCurr_Trx_Rt.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCurr_Trx_Rt.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurr_Trx_Rt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCurr_Trx_Rt.Location = New System.Drawing.Point(122, 32)
        Me.lblCurr_Trx_Rt.Name = "lblCurr_Trx_Rt"
        Me.lblCurr_Trx_Rt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCurr_Trx_Rt.Size = New System.Drawing.Size(120, 20)
        Me.lblCurr_Trx_Rt.TabIndex = 8
        Me.lblCurr_Trx_Rt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label32
        '
        Me.Label32.BackColor = System.Drawing.SystemColors.Control
        Me.Label32.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label32.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label32.Location = New System.Drawing.Point(6, 10)
        Me.Label32.Name = "Label32"
        Me.Label32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label32.Size = New System.Drawing.Size(113, 20)
        Me.Label32.TabIndex = 0
        Me.Label32.Text = "Total: Line Items"
        '
        'lblOrder_Item_Count
        '
        Me.lblOrder_Item_Count.BackColor = System.Drawing.SystemColors.Control
        Me.lblOrder_Item_Count.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOrder_Item_Count.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOrder_Item_Count.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOrder_Item_Count.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrder_Item_Count.Location = New System.Drawing.Point(122, 9)
        Me.lblOrder_Item_Count.Name = "lblOrder_Item_Count"
        Me.lblOrder_Item_Count.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOrder_Item_Count.Size = New System.Drawing.Size(120, 20)
        Me.lblOrder_Item_Count.TabIndex = 1
        Me.lblOrder_Item_Count.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ucOrderTotal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.lblShip_Amt)
        Me.Controls.Add(Me.lblOrig_Curr_Amt)
        Me.Controls.Add(Me.lblOrder_Amt)
        Me.Controls.Add(Me.lblTo_Curr_Cd_Desc)
        Me.Controls.Add(Me.lblFrom_Curr_Cd_Desc)
        Me.Controls.Add(Me.Label38)
        Me.Controls.Add(Me.Label37)
        Me.Controls.Add(Me.lblShip_Qty)
        Me.Controls.Add(Me.lblConv_Curr_Amt)
        Me.Controls.Add(Me.lblOrder_Qty)
        Me.Controls.Add(Me.lblCurr_Trx_Rt)
        Me.Controls.Add(Me.Label32)
        Me.Controls.Add(Me.lblOrder_Item_Count)
        Me.Name = "ucOrderTotal"
        Me.Size = New System.Drawing.Size(980, 555)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblShip_Amt As System.Windows.Forms.Label
    Friend WithEvents lblOrig_Curr_Amt As System.Windows.Forms.Label
    Friend WithEvents lblOrder_Amt As System.Windows.Forms.Label
    Friend WithEvents lblTo_Curr_Cd_Desc As System.Windows.Forms.Label
    Friend WithEvents lblFrom_Curr_Cd_Desc As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents lblShip_Qty As System.Windows.Forms.Label
    Friend WithEvents lblConv_Curr_Amt As System.Windows.Forms.Label
    Friend WithEvents lblOrder_Qty As System.Windows.Forms.Label
    Friend WithEvents lblCurr_Trx_Rt As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents lblOrder_Item_Count As System.Windows.Forms.Label

End Class
