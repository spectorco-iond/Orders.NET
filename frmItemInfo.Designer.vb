<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmItemInfo
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.lblMax_Reorder_Lvl = New System.Windows.Forms.Label
        Me.lblSum_Qty_In_Transit = New System.Windows.Forms.Label
        Me.lblSum_Qty_On_Ord_PO = New System.Windows.Forms.Label
        Me.lblSum_Qty_BkOrd = New System.Windows.Forms.Label
        Me.lblItem_No = New System.Windows.Forms.Label
        Me.lblSum_Qty_On_Ord = New System.Windows.Forms.Label
        Me.cmdClose = New System.Windows.Forms.Button
        Me.dgvLocDetails = New System.Windows.Forms.DataGridView
        Me.txtSum_Qty_On_Ord = New Orders.xTextBox
        Me.txtMax_Reorder_Lvl = New Orders.xTextBox
        Me.txtSum_Qty_In_Transit = New Orders.xTextBox
        Me.txtSum_Qty_On_Ord_PO = New Orders.xTextBox
        Me.txtSum_Qty_BkOrd = New Orders.xTextBox
        Me.txtItem_Desc_2 = New Orders.xTextBox
        Me.txtItem_Desc_1 = New Orders.xTextBox
        Me.txtItem_No = New Orders.xTextBox
        CType(Me.dgvLocDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblMax_Reorder_Lvl
        '
        Me.lblMax_Reorder_Lvl.BackColor = System.Drawing.SystemColors.Control
        Me.lblMax_Reorder_Lvl.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMax_Reorder_Lvl.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblMax_Reorder_Lvl.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMax_Reorder_Lvl.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblMax_Reorder_Lvl.Location = New System.Drawing.Point(10, 215)
        Me.lblMax_Reorder_Lvl.Name = "lblMax_Reorder_Lvl"
        Me.lblMax_Reorder_Lvl.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMax_Reorder_Lvl.Size = New System.Drawing.Size(185, 20)
        Me.lblMax_Reorder_Lvl.TabIndex = 53
        Me.lblMax_Reorder_Lvl.Text = "Reorder Level"
        '
        'lblSum_Qty_In_Transit
        '
        Me.lblSum_Qty_In_Transit.BackColor = System.Drawing.SystemColors.Control
        Me.lblSum_Qty_In_Transit.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSum_Qty_In_Transit.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblSum_Qty_In_Transit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSum_Qty_In_Transit.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblSum_Qty_In_Transit.Location = New System.Drawing.Point(10, 185)
        Me.lblSum_Qty_In_Transit.Name = "lblSum_Qty_In_Transit"
        Me.lblSum_Qty_In_Transit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSum_Qty_In_Transit.Size = New System.Drawing.Size(185, 20)
        Me.lblSum_Qty_In_Transit.TabIndex = 60
        Me.lblSum_Qty_In_Transit.Text = "Quantity Intransit"
        '
        'lblSum_Qty_On_Ord_PO
        '
        Me.lblSum_Qty_On_Ord_PO.BackColor = System.Drawing.SystemColors.Control
        Me.lblSum_Qty_On_Ord_PO.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSum_Qty_On_Ord_PO.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblSum_Qty_On_Ord_PO.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSum_Qty_On_Ord_PO.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblSum_Qty_On_Ord_PO.Location = New System.Drawing.Point(10, 155)
        Me.lblSum_Qty_On_Ord_PO.Name = "lblSum_Qty_On_Ord_PO"
        Me.lblSum_Qty_On_Ord_PO.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSum_Qty_On_Ord_PO.Size = New System.Drawing.Size(185, 20)
        Me.lblSum_Qty_On_Ord_PO.TabIndex = 58
        Me.lblSum_Qty_On_Ord_PO.Text = "Quantity on Order PO"
        '
        'lblSum_Qty_BkOrd
        '
        Me.lblSum_Qty_BkOrd.BackColor = System.Drawing.SystemColors.Control
        Me.lblSum_Qty_BkOrd.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSum_Qty_BkOrd.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblSum_Qty_BkOrd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSum_Qty_BkOrd.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblSum_Qty_BkOrd.Location = New System.Drawing.Point(10, 98)
        Me.lblSum_Qty_BkOrd.Name = "lblSum_Qty_BkOrd"
        Me.lblSum_Qty_BkOrd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSum_Qty_BkOrd.Size = New System.Drawing.Size(185, 20)
        Me.lblSum_Qty_BkOrd.TabIndex = 52
        Me.lblSum_Qty_BkOrd.Text = "Quantity Backordered"
        '
        'lblItem_No
        '
        Me.lblItem_No.BackColor = System.Drawing.SystemColors.Control
        Me.lblItem_No.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblItem_No.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblItem_No.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblItem_No.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblItem_No.Location = New System.Drawing.Point(10, 15)
        Me.lblItem_No.Name = "lblItem_No"
        Me.lblItem_No.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblItem_No.Size = New System.Drawing.Size(185, 20)
        Me.lblItem_No.TabIndex = 50
        Me.lblItem_No.Text = "Item No"
        '
        'lblSum_Qty_On_Ord
        '
        Me.lblSum_Qty_On_Ord.BackColor = System.Drawing.SystemColors.Control
        Me.lblSum_Qty_On_Ord.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSum_Qty_On_Ord.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblSum_Qty_On_Ord.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSum_Qty_On_Ord.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblSum_Qty_On_Ord.Location = New System.Drawing.Point(10, 127)
        Me.lblSum_Qty_On_Ord.Name = "lblSum_Qty_On_Ord"
        Me.lblSum_Qty_On_Ord.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSum_Qty_On_Ord.Size = New System.Drawing.Size(185, 20)
        Me.lblSum_Qty_On_Ord.TabIndex = 62
        Me.lblSum_Qty_On_Ord.Text = "Quantity on order"
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdClose.Location = New System.Drawing.Point(514, 544)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(75, 24)
        Me.cmdClose.TabIndex = 125
        Me.cmdClose.TabStop = False
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'dgvLocDetails
        '
        Me.dgvLocDetails.AllowUserToAddRows = False
        Me.dgvLocDetails.AllowUserToDeleteRows = False
        Me.dgvLocDetails.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvLocDetails.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvLocDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvLocDetails.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvLocDetails.Location = New System.Drawing.Point(13, 238)
        Me.dgvLocDetails.Name = "dgvLocDetails"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvLocDetails.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvLocDetails.Size = New System.Drawing.Size(576, 300)
        Me.dgvLocDetails.TabIndex = 126
        Me.dgvLocDetails.TabStop = False
        '
        'txtSum_Qty_On_Ord
        '
        Me.txtSum_Qty_On_Ord.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtSum_Qty_On_Ord.DataLength = CType(25, Long)
        Me.txtSum_Qty_On_Ord.DataType = Orders.CDataType.dtNumWithDecimals
        Me.txtSum_Qty_On_Ord.DateValue = New Date(CType(0, Long))
        Me.txtSum_Qty_On_Ord.DecimalDigits = CType(2, Long)
        Me.txtSum_Qty_On_Ord.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSum_Qty_On_Ord.Location = New System.Drawing.Point(201, 124)
        Me.txtSum_Qty_On_Ord.Mask = Nothing
        Me.txtSum_Qty_On_Ord.Name = "txtSum_Qty_On_Ord"
        Me.txtSum_Qty_On_Ord.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtSum_Qty_On_Ord.OldValue = Nothing
        Me.txtSum_Qty_On_Ord.ReadOnly = True
        Me.txtSum_Qty_On_Ord.Size = New System.Drawing.Size(214, 22)
        Me.txtSum_Qty_On_Ord.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtSum_Qty_On_Ord.StringValue = Nothing
        Me.txtSum_Qty_On_Ord.TabIndex = 63
        Me.txtSum_Qty_On_Ord.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSum_Qty_On_Ord.TextDataID = Nothing
        '
        'txtMax_Reorder_Lvl
        '
        Me.txtMax_Reorder_Lvl.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtMax_Reorder_Lvl.DataLength = CType(25, Long)
        Me.txtMax_Reorder_Lvl.DataType = Orders.CDataType.dtNumWithDecimals
        Me.txtMax_Reorder_Lvl.DateValue = New Date(CType(0, Long))
        Me.txtMax_Reorder_Lvl.DecimalDigits = CType(2, Long)
        Me.txtMax_Reorder_Lvl.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMax_Reorder_Lvl.Location = New System.Drawing.Point(201, 208)
        Me.txtMax_Reorder_Lvl.Mask = Nothing
        Me.txtMax_Reorder_Lvl.Name = "txtMax_Reorder_Lvl"
        Me.txtMax_Reorder_Lvl.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtMax_Reorder_Lvl.OldValue = Nothing
        Me.txtMax_Reorder_Lvl.ReadOnly = True
        Me.txtMax_Reorder_Lvl.Size = New System.Drawing.Size(214, 22)
        Me.txtMax_Reorder_Lvl.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtMax_Reorder_Lvl.StringValue = Nothing
        Me.txtMax_Reorder_Lvl.TabIndex = 55
        Me.txtMax_Reorder_Lvl.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtMax_Reorder_Lvl.TextDataID = Nothing
        '
        'txtSum_Qty_In_Transit
        '
        Me.txtSum_Qty_In_Transit.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtSum_Qty_In_Transit.DataLength = CType(25, Long)
        Me.txtSum_Qty_In_Transit.DataType = Orders.CDataType.dtNumWithDecimals
        Me.txtSum_Qty_In_Transit.DateValue = New Date(CType(0, Long))
        Me.txtSum_Qty_In_Transit.DecimalDigits = CType(2, Long)
        Me.txtSum_Qty_In_Transit.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSum_Qty_In_Transit.Location = New System.Drawing.Point(201, 180)
        Me.txtSum_Qty_In_Transit.Mask = Nothing
        Me.txtSum_Qty_In_Transit.Name = "txtSum_Qty_In_Transit"
        Me.txtSum_Qty_In_Transit.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtSum_Qty_In_Transit.OldValue = Nothing
        Me.txtSum_Qty_In_Transit.ReadOnly = True
        Me.txtSum_Qty_In_Transit.Size = New System.Drawing.Size(214, 22)
        Me.txtSum_Qty_In_Transit.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtSum_Qty_In_Transit.StringValue = Nothing
        Me.txtSum_Qty_In_Transit.TabIndex = 61
        Me.txtSum_Qty_In_Transit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSum_Qty_In_Transit.TextDataID = Nothing
        '
        'txtSum_Qty_On_Ord_PO
        '
        Me.txtSum_Qty_On_Ord_PO.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtSum_Qty_On_Ord_PO.DataLength = CType(25, Long)
        Me.txtSum_Qty_On_Ord_PO.DataType = Orders.CDataType.dtNumWithDecimals
        Me.txtSum_Qty_On_Ord_PO.DateValue = New Date(CType(0, Long))
        Me.txtSum_Qty_On_Ord_PO.DecimalDigits = CType(2, Long)
        Me.txtSum_Qty_On_Ord_PO.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSum_Qty_On_Ord_PO.Location = New System.Drawing.Point(201, 152)
        Me.txtSum_Qty_On_Ord_PO.Mask = Nothing
        Me.txtSum_Qty_On_Ord_PO.Name = "txtSum_Qty_On_Ord_PO"
        Me.txtSum_Qty_On_Ord_PO.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtSum_Qty_On_Ord_PO.OldValue = Nothing
        Me.txtSum_Qty_On_Ord_PO.ReadOnly = True
        Me.txtSum_Qty_On_Ord_PO.Size = New System.Drawing.Size(214, 22)
        Me.txtSum_Qty_On_Ord_PO.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtSum_Qty_On_Ord_PO.StringValue = Nothing
        Me.txtSum_Qty_On_Ord_PO.TabIndex = 59
        Me.txtSum_Qty_On_Ord_PO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSum_Qty_On_Ord_PO.TextDataID = Nothing
        '
        'txtSum_Qty_BkOrd
        '
        Me.txtSum_Qty_BkOrd.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtSum_Qty_BkOrd.DataLength = CType(25, Long)
        Me.txtSum_Qty_BkOrd.DataType = Orders.CDataType.dtNumWithDecimals
        Me.txtSum_Qty_BkOrd.DateValue = New Date(CType(0, Long))
        Me.txtSum_Qty_BkOrd.DecimalDigits = CType(2, Long)
        Me.txtSum_Qty_BkOrd.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSum_Qty_BkOrd.Location = New System.Drawing.Point(201, 96)
        Me.txtSum_Qty_BkOrd.Mask = Nothing
        Me.txtSum_Qty_BkOrd.Name = "txtSum_Qty_BkOrd"
        Me.txtSum_Qty_BkOrd.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtSum_Qty_BkOrd.OldValue = Nothing
        Me.txtSum_Qty_BkOrd.ReadOnly = True
        Me.txtSum_Qty_BkOrd.Size = New System.Drawing.Size(214, 22)
        Me.txtSum_Qty_BkOrd.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtSum_Qty_BkOrd.StringValue = Nothing
        Me.txtSum_Qty_BkOrd.TabIndex = 57
        Me.txtSum_Qty_BkOrd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSum_Qty_BkOrd.TextDataID = Nothing
        '
        'txtItem_Desc_2
        '
        Me.txtItem_Desc_2.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtItem_Desc_2.DataLength = CType(25, Long)
        Me.txtItem_Desc_2.DataType = Orders.CDataType.dtString
        Me.txtItem_Desc_2.DateValue = New Date(CType(0, Long))
        Me.txtItem_Desc_2.DecimalDigits = CType(0, Long)
        Me.txtItem_Desc_2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItem_Desc_2.Location = New System.Drawing.Point(201, 68)
        Me.txtItem_Desc_2.Mask = Nothing
        Me.txtItem_Desc_2.Name = "txtItem_Desc_2"
        Me.txtItem_Desc_2.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtItem_Desc_2.OldValue = Nothing
        Me.txtItem_Desc_2.ReadOnly = True
        Me.txtItem_Desc_2.Size = New System.Drawing.Size(214, 22)
        Me.txtItem_Desc_2.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtItem_Desc_2.StringValue = Nothing
        Me.txtItem_Desc_2.TabIndex = 56
        Me.txtItem_Desc_2.Text = "BARREL & MATCHING GRIP"
        Me.txtItem_Desc_2.TextDataID = Nothing
        '
        'txtItem_Desc_1
        '
        Me.txtItem_Desc_1.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtItem_Desc_1.DataLength = CType(25, Long)
        Me.txtItem_Desc_1.DataType = Orders.CDataType.dtString
        Me.txtItem_Desc_1.DateValue = New Date(CType(0, Long))
        Me.txtItem_Desc_1.DecimalDigits = CType(0, Long)
        Me.txtItem_Desc_1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItem_Desc_1.Location = New System.Drawing.Point(201, 40)
        Me.txtItem_Desc_1.Mask = Nothing
        Me.txtItem_Desc_1.Name = "txtItem_Desc_1"
        Me.txtItem_Desc_1.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtItem_Desc_1.OldValue = Nothing
        Me.txtItem_Desc_1.ReadOnly = True
        Me.txtItem_Desc_1.Size = New System.Drawing.Size(214, 22)
        Me.txtItem_Desc_1.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtItem_Desc_1.StringValue = Nothing
        Me.txtItem_Desc_1.TabIndex = 54
        Me.txtItem_Desc_1.TextDataID = Nothing
        '
        'txtItem_No
        '
        Me.txtItem_No.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtItem_No.DataLength = CType(25, Long)
        Me.txtItem_No.DataType = Orders.CDataType.dtString
        Me.txtItem_No.DateValue = New Date(CType(0, Long))
        Me.txtItem_No.DecimalDigits = CType(0, Long)
        Me.txtItem_No.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItem_No.Location = New System.Drawing.Point(201, 12)
        Me.txtItem_No.Mask = Nothing
        Me.txtItem_No.Name = "txtItem_No"
        Me.txtItem_No.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtItem_No.OldValue = Nothing
        Me.txtItem_No.ReadOnly = True
        Me.txtItem_No.Size = New System.Drawing.Size(214, 22)
        Me.txtItem_No.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtItem_No.StringValue = Nothing
        Me.txtItem_No.TabIndex = 51
        Me.txtItem_No.TextDataID = Nothing
        '
        'frmItemInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(601, 580)
        Me.Controls.Add(Me.dgvLocDetails)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.txtSum_Qty_On_Ord)
        Me.Controls.Add(Me.lblSum_Qty_On_Ord)
        Me.Controls.Add(Me.txtMax_Reorder_Lvl)
        Me.Controls.Add(Me.lblMax_Reorder_Lvl)
        Me.Controls.Add(Me.txtSum_Qty_In_Transit)
        Me.Controls.Add(Me.lblSum_Qty_In_Transit)
        Me.Controls.Add(Me.txtSum_Qty_On_Ord_PO)
        Me.Controls.Add(Me.lblSum_Qty_On_Ord_PO)
        Me.Controls.Add(Me.txtSum_Qty_BkOrd)
        Me.Controls.Add(Me.txtItem_Desc_2)
        Me.Controls.Add(Me.txtItem_Desc_1)
        Me.Controls.Add(Me.lblSum_Qty_BkOrd)
        Me.Controls.Add(Me.txtItem_No)
        Me.Controls.Add(Me.lblItem_No)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmItemInfo"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Item Location Detail View"
        CType(Me.dgvLocDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtMax_Reorder_Lvl As Orders.xTextBox
    Friend WithEvents lblMax_Reorder_Lvl As System.Windows.Forms.Label
    Friend WithEvents txtSum_Qty_In_Transit As Orders.xTextBox
    Friend WithEvents lblSum_Qty_In_Transit As System.Windows.Forms.Label
    Friend WithEvents txtSum_Qty_On_Ord_PO As Orders.xTextBox
    Friend WithEvents lblSum_Qty_On_Ord_PO As System.Windows.Forms.Label
    Friend WithEvents txtSum_Qty_BkOrd As Orders.xTextBox
    Friend WithEvents txtItem_Desc_2 As Orders.xTextBox
    Friend WithEvents txtItem_Desc_1 As Orders.xTextBox
    Friend WithEvents lblSum_Qty_BkOrd As System.Windows.Forms.Label
    Friend WithEvents txtItem_No As Orders.xTextBox
    Friend WithEvents lblItem_No As System.Windows.Forms.Label
    Friend WithEvents txtSum_Qty_On_Ord As Orders.xTextBox
    Friend WithEvents lblSum_Qty_On_Ord As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents dgvLocDetails As System.Windows.Forms.DataGridView
End Class
