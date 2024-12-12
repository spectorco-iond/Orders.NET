<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmSpecialPricing
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.lblPriceCodes = New System.Windows.Forms.Label()
        Me.lblSpecialPrices = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lblCustomer = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.fraSpecialPriceHistory = New System.Windows.Forms.GroupBox()
        Me.lblEndDateAfter = New System.Windows.Forms.Label()
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker()
        Me.optAllItems = New System.Windows.Forms.RadioButton()
        Me.optThisStyle = New System.Windows.Forms.RadioButton()
        Me.optThisItem = New System.Windows.Forms.RadioButton()
        Me.cmdViewSpecialPricesHistory = New System.Windows.Forms.Button()
        Me.txtProg_Comments = New Orders.xTextBox()
        Me.gridSpecialPricesHISTORY = New Orders.XDataGridView()
        Me.txtItemNo = New Orders.xTextBox()
        Me.txtCusNo = New Orders.xTextBox()
        Me.gridSpecialPrices = New Orders.XDataGridView()
        Me.gridPriceCodes = New Orders.XDataGridView()
        Me.fraSpecialPriceHistory.SuspendLayout()
        CType(Me.gridSpecialPricesHISTORY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gridSpecialPrices, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gridPriceCodes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblPriceCodes
        '
        Me.lblPriceCodes.AutoSize = True
        Me.lblPriceCodes.Location = New System.Drawing.Point(13, 42)
        Me.lblPriceCodes.Name = "lblPriceCodes"
        Me.lblPriceCodes.Size = New System.Drawing.Size(56, 13)
        Me.lblPriceCodes.TabIndex = 1
        Me.lblPriceCodes.Text = "Comments"
        '
        'lblSpecialPrices
        '
        Me.lblSpecialPrices.AutoSize = True
        Me.lblSpecialPrices.Location = New System.Drawing.Point(13, 197)
        Me.lblSpecialPrices.Name = "lblSpecialPrices"
        Me.lblSpecialPrices.Size = New System.Drawing.Size(74, 13)
        Me.lblSpecialPrices.TabIndex = 2
        Me.lblSpecialPrices.Text = "Special Prices"
        '
        'cmdCancel
        '
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.Location = New System.Drawing.Point(817, 353)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(100, 24)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'lblCustomer
        '
        Me.lblCustomer.BackColor = System.Drawing.SystemColors.Control
        Me.lblCustomer.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCustomer.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.lblCustomer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCustomer.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblCustomer.Location = New System.Drawing.Point(13, 9)
        Me.lblCustomer.Name = "lblCustomer"
        Me.lblCustomer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCustomer.Size = New System.Drawing.Size(85, 20)
        Me.lblCustomer.TabIndex = 27
        Me.lblCustomer.Text = "Customer"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label2.Location = New System.Drawing.Point(213, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(85, 20)
        Me.Label2.TabIndex = 31
        Me.Label2.Text = "Item No"
        '
        'fraSpecialPriceHistory
        '
        Me.fraSpecialPriceHistory.Controls.Add(Me.lblEndDateAfter)
        Me.fraSpecialPriceHistory.Controls.Add(Me.dtpEndDate)
        Me.fraSpecialPriceHistory.Controls.Add(Me.optAllItems)
        Me.fraSpecialPriceHistory.Controls.Add(Me.optThisStyle)
        Me.fraSpecialPriceHistory.Controls.Add(Me.optThisItem)
        Me.fraSpecialPriceHistory.Location = New System.Drawing.Point(16, 383)
        Me.fraSpecialPriceHistory.Name = "fraSpecialPriceHistory"
        Me.fraSpecialPriceHistory.Size = New System.Drawing.Size(513, 50)
        Me.fraSpecialPriceHistory.TabIndex = 34
        Me.fraSpecialPriceHistory.TabStop = False
        Me.fraSpecialPriceHistory.Text = "Special Price History"
        '
        'lblEndDateAfter
        '
        Me.lblEndDateAfter.AutoSize = True
        Me.lblEndDateAfter.Location = New System.Drawing.Point(288, 21)
        Me.lblEndDateAfter.Name = "lblEndDateAfter"
        Me.lblEndDateAfter.Size = New System.Drawing.Size(77, 13)
        Me.lblEndDateAfter.TabIndex = 39
        Me.lblEndDateAfter.Text = "End Date After"
        '
        'dtpEndDate
        '
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpEndDate.Location = New System.Drawing.Point(371, 19)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(127, 20)
        Me.dtpEndDate.TabIndex = 38
        '
        'optAllItems
        '
        Me.optAllItems.AutoSize = True
        Me.optAllItems.Location = New System.Drawing.Point(161, 19)
        Me.optAllItems.Name = "optAllItems"
        Me.optAllItems.Size = New System.Drawing.Size(64, 17)
        Me.optAllItems.TabIndex = 37
        Me.optAllItems.Text = "All Items"
        Me.optAllItems.UseVisualStyleBackColor = True
        '
        'optThisStyle
        '
        Me.optThisStyle.AutoSize = True
        Me.optThisStyle.Location = New System.Drawing.Point(84, 19)
        Me.optThisStyle.Name = "optThisStyle"
        Me.optThisStyle.Size = New System.Drawing.Size(71, 17)
        Me.optThisStyle.TabIndex = 36
        Me.optThisStyle.Text = "This Style"
        Me.optThisStyle.UseVisualStyleBackColor = True
        '
        'optThisItem
        '
        Me.optThisItem.AutoSize = True
        Me.optThisItem.Checked = True
        Me.optThisItem.Location = New System.Drawing.Point(10, 19)
        Me.optThisItem.Name = "optThisItem"
        Me.optThisItem.Size = New System.Drawing.Size(68, 17)
        Me.optThisItem.TabIndex = 35
        Me.optThisItem.TabStop = True
        Me.optThisItem.Text = "This Item"
        Me.optThisItem.UseVisualStyleBackColor = True
        '
        'cmdViewSpecialPricesHistory
        '
        Me.cmdViewSpecialPricesHistory.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdViewSpecialPricesHistory.Location = New System.Drawing.Point(13, 353)
        Me.cmdViewSpecialPricesHistory.Name = "cmdViewSpecialPricesHistory"
        Me.cmdViewSpecialPricesHistory.Size = New System.Drawing.Size(153, 24)
        Me.cmdViewSpecialPricesHistory.TabIndex = 35
        Me.cmdViewSpecialPricesHistory.Text = "Special Price History"
        Me.cmdViewSpecialPricesHistory.UseVisualStyleBackColor = True
        '
        'txtProg_Comments
        '
        Me.txtProg_Comments.CharacterInput = Orders.Cinput.CharactersOnly
        Me.txtProg_Comments.DataLength = CType(1000, Long)
        Me.txtProg_Comments.DataType = Orders.CDataType.dtString
        Me.txtProg_Comments.DateValue = New Date(CType(0, Long))
        Me.txtProg_Comments.DecimalDigits = CType(0, Long)
        Me.txtProg_Comments.Font = New System.Drawing.Font("Arial", 10.5!)
        Me.txtProg_Comments.Location = New System.Drawing.Point(12, 58)
        Me.txtProg_Comments.Mask = Nothing
        Me.txtProg_Comments.Multiline = True
        Me.txtProg_Comments.Name = "txtProg_Comments"
        Me.txtProg_Comments.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtProg_Comments.OldValue = Nothing
        Me.txtProg_Comments.ReadOnly = True
        Me.txtProg_Comments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtProg_Comments.Size = New System.Drawing.Size(905, 119)
        Me.txtProg_Comments.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtProg_Comments.StringValue = Nothing
        Me.txtProg_Comments.TabIndex = 37
        Me.txtProg_Comments.TextDataID = Nothing
        '
        'gridSpecialPricesHISTORY
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.gridSpecialPricesHISTORY.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.gridSpecialPricesHISTORY.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.gridSpecialPricesHISTORY.DefaultCellStyle = DataGridViewCellStyle2
        Me.gridSpecialPricesHISTORY.Location = New System.Drawing.Point(12, 439)
        Me.gridSpecialPricesHISTORY.Name = "gridSpecialPricesHISTORY"
        Me.gridSpecialPricesHISTORY.Size = New System.Drawing.Size(904, 134)
        Me.gridSpecialPricesHISTORY.TabIndex = 36
        '
        'txtItemNo
        '
        Me.txtItemNo.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtItemNo.DataLength = CType(20, Long)
        Me.txtItemNo.DataType = Orders.CDataType.dtString
        Me.txtItemNo.DateValue = New Date(CType(0, Long))
        Me.txtItemNo.DecimalDigits = CType(0, Long)
        Me.txtItemNo.Font = New System.Drawing.Font("Arial", 10.5!)
        Me.txtItemNo.Location = New System.Drawing.Point(274, 4)
        Me.txtItemNo.Mask = Nothing
        Me.txtItemNo.Name = "txtItemNo"
        Me.txtItemNo.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtItemNo.OldValue = Nothing
        Me.txtItemNo.Size = New System.Drawing.Size(105, 24)
        Me.txtItemNo.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtItemNo.StringValue = Nothing
        Me.txtItemNo.TabIndex = 32
        Me.txtItemNo.TextDataID = Nothing
        '
        'txtCusNo
        '
        Me.txtCusNo.CharacterInput = Orders.Cinput.NumericOnly
        Me.txtCusNo.DataLength = CType(20, Long)
        Me.txtCusNo.DataType = Orders.CDataType.dtString
        Me.txtCusNo.DateValue = New Date(CType(0, Long))
        Me.txtCusNo.DecimalDigits = CType(0, Long)
        Me.txtCusNo.Font = New System.Drawing.Font("Arial", 10.5!)
        Me.txtCusNo.Location = New System.Drawing.Point(82, 5)
        Me.txtCusNo.Mask = Nothing
        Me.txtCusNo.Name = "txtCusNo"
        Me.txtCusNo.NumericValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtCusNo.OldValue = Nothing
        Me.txtCusNo.Size = New System.Drawing.Size(105, 24)
        Me.txtCusNo.SpacePadding = Orders.CSpacePadding.NoPadding
        Me.txtCusNo.StringValue = Nothing
        Me.txtCusNo.TabIndex = 28
        Me.txtCusNo.TextDataID = Nothing
        '
        'gridSpecialPrices
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.gridSpecialPrices.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.gridSpecialPrices.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.gridSpecialPrices.DefaultCellStyle = DataGridViewCellStyle4
        Me.gridSpecialPrices.Location = New System.Drawing.Point(13, 213)
        Me.gridSpecialPrices.Name = "gridSpecialPrices"
        Me.gridSpecialPrices.Size = New System.Drawing.Size(904, 134)
        Me.gridSpecialPrices.TabIndex = 3
        '
        'gridPriceCodes
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.gridPriceCodes.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.gridPriceCodes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.gridPriceCodes.DefaultCellStyle = DataGridViewCellStyle6
        Me.gridPriceCodes.Enabled = False
        Me.gridPriceCodes.Location = New System.Drawing.Point(13, 58)
        Me.gridPriceCodes.Name = "gridPriceCodes"
        Me.gridPriceCodes.Size = New System.Drawing.Size(904, 119)
        Me.gridPriceCodes.TabIndex = 0
        Me.gridPriceCodes.Visible = False
        '
        'frmSpecialPricing
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(975, 580)
        Me.Controls.Add(Me.txtProg_Comments)
        Me.Controls.Add(Me.gridSpecialPricesHISTORY)
        Me.Controls.Add(Me.cmdViewSpecialPricesHistory)
        Me.Controls.Add(Me.fraSpecialPriceHistory)
        Me.Controls.Add(Me.txtItemNo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtCusNo)
        Me.Controls.Add(Me.lblCustomer)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.gridSpecialPrices)
        Me.Controls.Add(Me.lblSpecialPrices)
        Me.Controls.Add(Me.lblPriceCodes)
        Me.Controls.Add(Me.gridPriceCodes)
        Me.Name = "frmSpecialPricing"
        Me.Text = "Special Pricing"
        Me.fraSpecialPriceHistory.ResumeLayout(False)
        Me.fraSpecialPriceHistory.PerformLayout()
        CType(Me.gridSpecialPricesHISTORY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gridSpecialPrices, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gridPriceCodes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents gridPriceCodes As XDataGridView
    Friend WithEvents lblPriceCodes As Label
    Friend WithEvents lblSpecialPrices As Label
    Friend WithEvents gridSpecialPrices As XDataGridView
    Friend WithEvents cmdCancel As Button
    Friend WithEvents txtCusNo As xTextBox
    Friend WithEvents lblCustomer As Label
    Friend WithEvents txtItemNo As xTextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents fraSpecialPriceHistory As GroupBox
    Friend WithEvents lblEndDateAfter As Label
    Friend WithEvents dtpEndDate As DateTimePicker
    Friend WithEvents optAllItems As RadioButton
    Friend WithEvents optThisStyle As RadioButton
    Friend WithEvents optThisItem As RadioButton
    Friend WithEvents cmdViewSpecialPricesHistory As Button
    Friend WithEvents gridSpecialPricesHISTORY As XDataGridView
    Friend WithEvents txtProg_Comments As xTextBox
End Class
